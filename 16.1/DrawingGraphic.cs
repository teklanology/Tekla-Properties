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
    public class DrawingGraphic
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private DrawingColors lineColour;
        private LineTypeEnum lineType;

        [Category("Line")]
        public DrawingColors LineColour
        {
            get { return lineColour; }
            set { lineColour = value; }
        }

        [Category("Line")]
        public LineTypeEnum LineType
        {
            get { return lineType; }
            set { lineType = value; }
        }

        public class Open : DrawingGraphic
        {
            private ArrowheadPositionEnum arrowheadPosition;
            private string arrowheadWidth;
            private string arrowheadHeight;

            [Category("Arrow")]
            public ArrowheadPositionEnum ArrowheadPosition
            {
                get { return arrowheadPosition; }
                set { arrowheadPosition = value; }
            }

            [Category("Arrow")]
            public string ArrowheadWidth
            {
                get { return arrowheadWidth; }
                set { arrowheadWidth = value; }
            }

            [Category("Arrow")]
            public string ArrowheadHeight
            {
                get { return arrowheadHeight; }
                set { arrowheadHeight = value; }
            }

            public class Line : Open
            {
                private string bulge;
                private Tekla.Structures.Geometry3d.Point startPoint;
                private Tekla.Structures.Geometry3d.Point endPoint;

                [Category("Line")]
                public string Bulge
                {
                    get { return bulge; }
                    set { bulge = value; }
                }

                public Tekla.Structures.Geometry3d.Point StartPoint
                {
                    get { return startPoint; }
                    set { startPoint = value; }
                }

                public Tekla.Structures.Geometry3d.Point EndPoint
                {
                    get { return endPoint; }
                    set { endPoint = value; }
                }

                /// <summary>
                /// Get Line
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Line)
                            {
                                Tekla.Structures.Drawing.Line line = (Tekla.Structures.Drawing.Line)drawingObjectEnum.Current;

                                if (line.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Both)
                                    arrowheadPosition = ArrowheadPositionEnum.Both;
                                else if (line.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.End)
                                    arrowheadPosition = ArrowheadPositionEnum.End;
                                else if (line.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.None)
                                    arrowheadPosition = ArrowheadPositionEnum.None;
                                else if (line.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Start)
                                    arrowheadPosition = ArrowheadPositionEnum.Start;

                                arrowheadHeight = line.Attributes.Arrowhead.Height.ToString();
                                arrowheadWidth = line.Attributes.Arrowhead.Width.ToString();
                                lineColour = line.Attributes.Line.Color;

                                if (line.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (line.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (line.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (line.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (line.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (line.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (line.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (line.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                bulge = line.Bulge.ToString();
                                endPoint = line.EndPoint;
                                startPoint = line.StartPoint;
                            }
                        }
                    }
                    if (drawingObjectEnum.GetSize() > 1)
                    {
                        arrowheadPosition = new ArrowheadPositionEnum();
                        arrowheadHeight = "";
                        arrowheadWidth = "";
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        bulge = "";
                        endPoint = null;
                        startPoint = null;
                    }
                }

                /// <summary>
                /// Modify Line
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    string label = e.ChangedItem.Label;
                    Drawing drawing = drawingHandler.GetActiveDrawing();
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    while (drawingObjectEnum.MoveNext())
                    {
                        if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Line)
                        {
                            Tekla.Structures.Drawing.Line line = (Tekla.Structures.Drawing.Line)drawingObjectEnum.Current;
                            if (label == "Bulge") line.Bulge = double.Parse(bulge);

                            if (label == "ArrowheadPosition")
                            {
                                if (arrowheadPosition == ArrowheadPositionEnum.Both)
                                    line.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Both;
                                else if (arrowheadPosition == ArrowheadPositionEnum.End)
                                    line.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.End;
                                else if (arrowheadPosition == ArrowheadPositionEnum.None)
                                    line.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.None;
                                else if (arrowheadPosition == ArrowheadPositionEnum.Start)
                                    line.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Start;
                            }

                            if (label == "ArrowheadHeight") line.Attributes.Arrowhead.Height = double.Parse(arrowheadHeight);
                            if (label == "ArrowheadWidth") line.Attributes.Arrowhead.Width = double.Parse(arrowheadWidth);
                            if (label == "LineColour") line.Attributes.Line.Color = lineColour;
                            if (label == "LineType")
                            {
                                if (lineType == LineTypeEnum.DashDot) line.Attributes.Line.Type = LineTypes.DashDot;
                                else if (lineType == LineTypeEnum.DashDoubleDot) line.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                else if (lineType == LineTypeEnum.DashedLine) line.Attributes.Line.Type = LineTypes.DashedLine;
                                else if (lineType == LineTypeEnum.DottedLine) line.Attributes.Line.Type = LineTypes.DottedLine;
                                else if (lineType == LineTypeEnum.SlashDash) line.Attributes.Line.Type = LineTypes.SlashDash;
                                else if (lineType == LineTypeEnum.SlashedLine) line.Attributes.Line.Type = LineTypes.SlashedLine;
                                else if (lineType == LineTypeEnum.SolidLine) line.Attributes.Line.Type = LineTypes.SolidLine;
                                else if (lineType == LineTypeEnum.UndefinedLine) line.Attributes.Line.Type = LineTypes.UndefinedLine;
                            }

                            line.Modify();
                            drawing.CommitChanges();
                        }
                    }
                }

            }

            public class Arc : Open
            {
                private string radius;
                private Tekla.Structures.Geometry3d.Point startPoint;
                private Tekla.Structures.Geometry3d.Point endPoint;

                public string Radius
                {
                    get { return radius; }
                    set { radius = value; }
                }

                public Tekla.Structures.Geometry3d.Point StartPoint
                {
                    get { return startPoint; }
                    set { startPoint = value; }
                }

                public Tekla.Structures.Geometry3d.Point EndPoint
                {
                    get { return endPoint; }
                    set { endPoint = value; }
                }

                /// <summary>
                /// Get Arc
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Arc)
                            {
                                Tekla.Structures.Drawing.Arc arc = (Tekla.Structures.Drawing.Arc)drawingObjectEnum.Current;
                                radius = arc.Radius.ToString("F02");
                                startPoint = arc.StartPoint;
                                endPoint = arc.EndPoint;
                                lineColour = arc.Attributes.Line.Color;

                                if (arc.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (arc.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (arc.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (arc.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (arc.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (arc.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (arc.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (arc.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                if (arc.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Both)
                                    arrowheadPosition = ArrowheadPositionEnum.Both;
                                else if (arc.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.End)
                                    arrowheadPosition = ArrowheadPositionEnum.End;
                                else if (arc.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.None)
                                    arrowheadPosition = ArrowheadPositionEnum.None;
                                else if (arc.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Start)
                                    arrowheadPosition = ArrowheadPositionEnum.Start;

                                arrowheadHeight = arc.Attributes.Arrowhead.Height.ToString();
                                arrowheadWidth = arc.Attributes.Arrowhead.Width.ToString();
                            }
                        }
                    }
                    if (drawingObjectEnum.GetSize() > 1)
                    {
                        radius = "";
                        startPoint = null;
                        endPoint = null;
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        arrowheadPosition = new ArrowheadPositionEnum();
                        arrowheadHeight = "";
                        arrowheadWidth = "";
                    }
                }

                /// <summary>
                /// Modify Arc
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    try
                    {
                        string label = e.ChangedItem.Label;
                        Drawing drawing = drawingHandler.GetActiveDrawing();
                        DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Arc)
                            {
                                Tekla.Structures.Drawing.Arc arc = (Tekla.Structures.Drawing.Arc)drawingObjectEnum.Current;
                                if (label == "Radius") arc.Radius = double.Parse(radius);
                                if (label == "LineColour") arc.Attributes.Line.Color = lineColour;
                                if (label == "LineType")
                                {
                                    if (lineType == LineTypeEnum.DashDot) arc.Attributes.Line.Type = LineTypes.DashDot;
                                    else if (lineType == LineTypeEnum.DashDoubleDot) arc.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                    else if (lineType == LineTypeEnum.DashedLine) arc.Attributes.Line.Type = LineTypes.DashedLine;
                                    else if (lineType == LineTypeEnum.DottedLine) arc.Attributes.Line.Type = LineTypes.DottedLine;
                                    else if (lineType == LineTypeEnum.SlashDash) arc.Attributes.Line.Type = LineTypes.SlashDash;
                                    else if (lineType == LineTypeEnum.SlashedLine) arc.Attributes.Line.Type = LineTypes.SlashedLine;
                                    else if (lineType == LineTypeEnum.SolidLine) arc.Attributes.Line.Type = LineTypes.SolidLine;
                                    else if (lineType == LineTypeEnum.UndefinedLine) arc.Attributes.Line.Type = LineTypes.UndefinedLine;
                                }

                                if (label == "ArrowheadPosition")
                                {
                                    if (arrowheadPosition == ArrowheadPositionEnum.Both)
                                        arc.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Both;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.End)
                                        arc.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.End;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.None)
                                        arc.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.None;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.Start)
                                        arc.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Start;
                                }

                                if (label == "ArrowheadHeight") arc.Attributes.Arrowhead.Height = double.Parse(arrowheadHeight);
                                if (label == "ArrowheadWidth") arc.Attributes.Arrowhead.Width = double.Parse(arrowheadWidth);
                                arc.Modify();
                                drawing.CommitChanges();
                            }
                        }
                    }
                    catch { }
                }

            }

            public class Polyline : Open
            {
                private string bulge;

                [Category("Line")]
                public string Bulge
                {
                    get { return bulge; }
                    set { bulge = value; }
                }

                /// <summary>
                /// Get Polyline
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Polyline)
                            {
                                Tekla.Structures.Drawing.Polyline polyline = (Tekla.Structures.Drawing.Polyline)drawingObjectEnum.Current;

                                if (polyline.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Both)
                                    arrowheadPosition = ArrowheadPositionEnum.Both;
                                else if (polyline.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.End)
                                    arrowheadPosition = ArrowheadPositionEnum.End;
                                else if (polyline.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.None)
                                    arrowheadPosition = ArrowheadPositionEnum.None;
                                else if (polyline.Attributes.Arrowhead.ArrowPosition == ArrowheadPositions.Start)
                                    arrowheadPosition = ArrowheadPositionEnum.Start;

                                arrowheadHeight = polyline.Attributes.Arrowhead.Height.ToString();
                                arrowheadWidth = polyline.Attributes.Arrowhead.Width.ToString();
                                lineColour = polyline.Attributes.Line.Color;

                                if (polyline.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (polyline.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (polyline.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (polyline.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (polyline.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (polyline.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (polyline.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (polyline.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                bulge = polyline.Bulge.ToString();
                            }
                        }
                    }
                    else if (drawingObjectEnum.GetSize() > 1)
                    {
                        arrowheadPosition = new ArrowheadPositionEnum();
                        arrowheadHeight = "";
                        arrowheadWidth = "";
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        bulge = "";
                    }
                }

                /// <summary>
                /// Modify Polyline
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    try
                    {
                        string label = e.ChangedItem.Label;
                        Drawing drawing = drawingHandler.GetActiveDrawing();
                        DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Polyline)
                            {
                                Tekla.Structures.Drawing.Polyline polyline = (Tekla.Structures.Drawing.Polyline)drawingObjectEnum.Current;
                                if (label == "Bulge") polyline.Bulge = double.Parse(bulge);

                                if (label == "ArrowheadPosition")
                                {
                                    if (arrowheadPosition == ArrowheadPositionEnum.Both)
                                        polyline.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Both;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.End)
                                        polyline.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.End;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.None)
                                        polyline.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.None;
                                    else if (arrowheadPosition == ArrowheadPositionEnum.Start)
                                        polyline.Attributes.Arrowhead.ArrowPosition = ArrowheadPositions.Start;
                                }

                                if (label == "ArrowheadHeight") polyline.Attributes.Arrowhead.Height = double.Parse(arrowheadHeight);
                                if (label == "ArrowheadWidth") polyline.Attributes.Arrowhead.Width = double.Parse(arrowheadWidth);
                                if (label == "LineColour") polyline.Attributes.Line.Color = lineColour;
                                if (label == "LineType")
                                {
                                    if (lineType == LineTypeEnum.DashDot) polyline.Attributes.Line.Type = LineTypes.DashDot;
                                    else if (lineType == LineTypeEnum.DashDoubleDot) polyline.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                    else if (lineType == LineTypeEnum.DashedLine) polyline.Attributes.Line.Type = LineTypes.DashedLine;
                                    else if (lineType == LineTypeEnum.DottedLine) polyline.Attributes.Line.Type = LineTypes.DottedLine;
                                    else if (lineType == LineTypeEnum.SlashDash) polyline.Attributes.Line.Type = LineTypes.SlashDash;
                                    else if (lineType == LineTypeEnum.SlashedLine) polyline.Attributes.Line.Type = LineTypes.SlashedLine;
                                    else if (lineType == LineTypeEnum.SolidLine) polyline.Attributes.Line.Type = LineTypes.SolidLine;
                                    else if (lineType == LineTypeEnum.UndefinedLine) polyline.Attributes.Line.Type = LineTypes.UndefinedLine;
                                }

                                polyline.Modify();
                                drawing.CommitChanges();
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        public class Closed : DrawingGraphic
        {
            private string hatchName;
            private string hatchAngle;
            private DrawingColors hatchColour;
            private DrawingColors hatchBackgroundColour;
            private Bool hatchDrawBackgroundColour;
            private Bool hatchFactorType;
            private string hatchOffsetX;
            private string hatchOffsetY;
            private string hatchScaleX;
            private string hatchScaleY;

            [Category("Fill")]
            public string HatchName
            {
                get { return hatchName; }
                set { hatchName = value; }
            }

            [Category("Fill")]
            public string HatchAngle
            {
                get { return hatchAngle; }
                set { hatchAngle = value; }
            }

            [Category("Fill")]
            public DrawingColors HatchColour
            {
                get { return hatchColour; }
                set { hatchColour = value; }
            }

            [Category("Fill")]
            public DrawingColors HatchBackgroundColour
            {
                get { return hatchBackgroundColour; }
                set { hatchBackgroundColour = value; }
            }

            [Category("Fill")]
            public Bool HatchDrawBackgroundColour
            {
                get { return hatchDrawBackgroundColour; }
                set { hatchDrawBackgroundColour = value; }
            }

            [Category("Fill")]
            public Bool HatchFactorType
            {
                get { return hatchFactorType; }
                set { hatchFactorType = value; }
            }

            [Category("Fill")]
            public string HatchOffsetX
            {
                get { return hatchOffsetX; }
                set { hatchOffsetX = value; }
            }

            [Category("Fill")]
            public string HatchOffsetY
            {
                get { return hatchOffsetY; }
                set { hatchOffsetY = value; }
            }

            [Category("Fill")]
            public string HatchScaleX
            {
                get { return hatchScaleX; }
                set { hatchScaleX = value; }
            }

            [Category("Fill")]
            public string HatchScaleY
            {
                get { return hatchScaleY; }
                set { hatchScaleY = value; }
            }

            public class Rectangle : Closed
            {
                private string bulge;
                private string width;
                private string height;

                [Category("Line")]
                public string Bulge
                {
                    get { return bulge; }
                    set { bulge = value; }
                }

                public string Width
                {
                    get { return width; }
                    set { width = value; }
                }

                public string Height
                {
                    get { return height; }
                    set { height = value; }
                }

                /// <summary>
                /// Get Rectangle
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Rectangle)
                            {
                                Tekla.Structures.Drawing.Rectangle rectangle = (Tekla.Structures.Drawing.Rectangle)drawingObjectEnum.Current;

                                bulge = rectangle.Attributes.Bulge.ToString();
                                lineColour = rectangle.Attributes.Line.Color;

                                if (rectangle.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (rectangle.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (rectangle.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (rectangle.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (rectangle.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (rectangle.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (rectangle.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (rectangle.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                hatchName = rectangle.Attributes.Hatch.Name;
                                hatchAngle = rectangle.Attributes.Hatch.Angle.ToString();
                                hatchBackgroundColour = rectangle.Attributes.Hatch.BackgroundColor;
                                hatchColour = rectangle.Attributes.Hatch.Color;
                                if (rectangle.Attributes.Hatch.DrawBackgroundColor) hatchDrawBackgroundColour = Bool.True;
                                else hatchDrawBackgroundColour = Bool.False;

                                if (rectangle.Attributes.Hatch.FactorType == 1) hatchFactorType = Bool.True; else hatchFactorType = Bool.False;
                                hatchOffsetX = rectangle.Attributes.Hatch.OffsetX.ToString();
                                hatchOffsetY = rectangle.Attributes.Hatch.OffsetY.ToString();
                                hatchScaleX = rectangle.Attributes.Hatch.ScaleX.ToString();
                                hatchScaleY = rectangle.Attributes.Hatch.ScaleY.ToString();

                                width = rectangle.Width.ToString("F02");
                                height = rectangle.Height.ToString("F02");
                            }
                        }
                    }
                    else if (drawingObjectEnum.GetSize() > 1)
                    {
                        bulge = "";
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        hatchName = "";
                        hatchAngle = "";
                        hatchBackgroundColour = new DrawingColors();
                        hatchColour = new DrawingColors();
                        hatchDrawBackgroundColour = new Bool();
                        hatchFactorType = new Bool();
                        hatchOffsetX = "";
                        hatchOffsetY = "";
                        hatchScaleX = "";
                        hatchScaleY = "";
                        width = "";
                        height = "";
                    }
                }

                /// <summary>
                /// Modify Rectangle
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    try
                    {
                        string label = e.ChangedItem.Label;
                        Drawing drawing = drawingHandler.GetActiveDrawing();
                        DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Rectangle)
                            {
                                Tekla.Structures.Drawing.Rectangle rectangle = (Tekla.Structures.Drawing.Rectangle)drawingObjectEnum.Current;

                                if (label == "Bulge") rectangle.Attributes.Bulge = double.Parse(bulge);
                                if (label == "LineColour") rectangle.Attributes.Line.Color = lineColour;
                                if (label == "LineType")
                                {
                                    if (lineType == LineTypeEnum.DashDot) rectangle.Attributes.Line.Type = LineTypes.DashDot;
                                    else if (lineType == LineTypeEnum.DashDoubleDot) rectangle.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                    else if (lineType == LineTypeEnum.DashedLine) rectangle.Attributes.Line.Type = LineTypes.DashedLine;
                                    else if (lineType == LineTypeEnum.DottedLine) rectangle.Attributes.Line.Type = LineTypes.DottedLine;
                                    else if (lineType == LineTypeEnum.SlashDash) rectangle.Attributes.Line.Type = LineTypes.SlashDash;
                                    else if (lineType == LineTypeEnum.SlashedLine) rectangle.Attributes.Line.Type = LineTypes.SlashedLine;
                                    else if (lineType == LineTypeEnum.SolidLine) rectangle.Attributes.Line.Type = LineTypes.SolidLine;
                                    else if (lineType == LineTypeEnum.UndefinedLine) rectangle.Attributes.Line.Type = LineTypes.UndefinedLine;
                                }

                                if (label == "Height") rectangle.Height = double.Parse(height);
                                if (label == "Width") rectangle.Width = double.Parse(width);


                                if (label == "HatchAngle") rectangle.Attributes.Hatch.Angle = double.Parse(hatchAngle);
                                if (label == "HatchBackgroundColour") rectangle.Attributes.Hatch.BackgroundColor = hatchBackgroundColour;
                                if (label == "HatchColour") rectangle.Attributes.Hatch.Color = hatchColour;
                                if (label == "HatchDrawBackgroundColour")
                                {
                                    if (hatchDrawBackgroundColour == Bool.True) rectangle.Attributes.Hatch.DrawBackgroundColor = true;
                                    else rectangle.Attributes.Hatch.DrawBackgroundColor = false;
                                }
                                if (label == "HatchFactorType")
                                {
                                    if (hatchFactorType == Bool.True) rectangle.Attributes.Hatch.FactorType = 1;
                                    else rectangle.Attributes.Hatch.FactorType = 0;
                                }

                                if (label == "HatchName") rectangle.Attributes.Hatch.Name = hatchName;
                                if (label == "HatchOffsetX") rectangle.Attributes.Hatch.OffsetX = double.Parse(hatchOffsetX);
                                if (label == "HatchOffsetY") rectangle.Attributes.Hatch.OffsetY = double.Parse(hatchOffsetY);
                                if (label == "HatchScaleX") rectangle.Attributes.Hatch.ScaleX = double.Parse(hatchScaleX);
                                if (label == "HatchScaleY") rectangle.Attributes.Hatch.ScaleY = double.Parse(hatchScaleY);

                                rectangle.Modify();
                                drawing.CommitChanges();
                            }
                        }
                    }
                    catch { }
                }
            }

            public class Circle : Closed
            {
                private string radius;
                private Tekla.Structures.Geometry3d.Point centerPoint;

                public string Radius
                {
                    get { return radius; }
                    set { radius = value; }
                }

                public Tekla.Structures.Geometry3d.Point CenterPoint
                {
                    get { return centerPoint; }
                    set { centerPoint = value; }
                }

                /// <summary>
                /// Get Circle
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Circle)
                            {
                                Tekla.Structures.Drawing.Circle circle = (Tekla.Structures.Drawing.Circle)drawingObjectEnum.Current;

                                lineColour = circle.Attributes.Line.Color;

                                if (circle.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (circle.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (circle.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (circle.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (circle.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (circle.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (circle.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (circle.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                radius = circle.Radius.ToString();

                                centerPoint = circle.CenterPoint;

                                hatchName = circle.Attributes.Hatch.Name;
                                hatchAngle = circle.Attributes.Hatch.Angle.ToString();
                                hatchBackgroundColour = circle.Attributes.Hatch.BackgroundColor;
                                hatchColour = circle.Attributes.Hatch.Color;
                                if (circle.Attributes.Hatch.DrawBackgroundColor) hatchDrawBackgroundColour = Bool.True;
                                else hatchDrawBackgroundColour = Bool.False;

                                if (circle.Attributes.Hatch.FactorType == 1) hatchFactorType = Bool.True; 
                                else hatchFactorType = Bool.False;
                                
                                hatchOffsetX = circle.Attributes.Hatch.OffsetX.ToString();
                                hatchOffsetY = circle.Attributes.Hatch.OffsetY.ToString();
                                hatchScaleX = circle.Attributes.Hatch.ScaleX.ToString();
                                hatchScaleY = circle.Attributes.Hatch.ScaleY.ToString();
                            }
                        }
                    }
                    if (drawingObjectEnum.GetSize() > 1)
                    {
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        radius = "";
                        centerPoint = null;
                        hatchName = "";
                        hatchAngle = "";
                        hatchBackgroundColour = new DrawingColors();
                        hatchColour = new DrawingColors();
                        hatchDrawBackgroundColour = new Bool();
                        hatchFactorType = new Bool();
                        hatchOffsetX = "";
                        hatchOffsetY = "";
                        hatchScaleX = "";
                        hatchScaleY = "";
                    }
                }

                /// <summary>
                /// Modify Circle
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    try
                    {
                        string label = e.ChangedItem.Label;
                        Drawing drawing = drawingHandler.GetActiveDrawing();
                        DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Circle)
                            {
                                Tekla.Structures.Drawing.Circle circle = (Tekla.Structures.Drawing.Circle)drawingObjectEnum.Current;
                                if (label == "Radius") circle.Radius = double.Parse(radius);

                                if (label == "LineColour") circle.Attributes.Line.Color = lineColour;
                                if (label == "LineType")
                                {
                                    if (lineType == LineTypeEnum.DashDot) circle.Attributes.Line.Type = LineTypes.DashDot;
                                    else if (lineType == LineTypeEnum.DashDoubleDot) circle.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                    else if (lineType == LineTypeEnum.DashedLine) circle.Attributes.Line.Type = LineTypes.DashedLine;
                                    else if (lineType == LineTypeEnum.DottedLine) circle.Attributes.Line.Type = LineTypes.DottedLine;
                                    else if (lineType == LineTypeEnum.SlashDash) circle.Attributes.Line.Type = LineTypes.SlashDash;
                                    else if (lineType == LineTypeEnum.SlashedLine) circle.Attributes.Line.Type = LineTypes.SlashedLine;
                                    else if (lineType == LineTypeEnum.SolidLine) circle.Attributes.Line.Type = LineTypes.SolidLine;
                                    else if (lineType == LineTypeEnum.UndefinedLine) circle.Attributes.Line.Type = LineTypes.UndefinedLine;
                                }

                                if (label == "HatchAngle") circle.Attributes.Hatch.Angle = double.Parse(hatchAngle);
                                if (label == "HatchBackgroundColour") circle.Attributes.Hatch.BackgroundColor = hatchBackgroundColour;
                                if (label == "HatchColour") circle.Attributes.Hatch.Color = hatchColour;

                                if (label == "HatchDrawBackgroundColour")
                                {
                                    if (hatchDrawBackgroundColour == Bool.True) circle.Attributes.Hatch.DrawBackgroundColor = true;
                                    else circle.Attributes.Hatch.DrawBackgroundColor = false;
                                }

                                if (label == "HatchFactorType")
                                {
                                    if (hatchFactorType == Bool.True) circle.Attributes.Hatch.FactorType = 1;
                                    else circle.Attributes.Hatch.FactorType = 0;
                                }

                                if (label == "HatchName") circle.Attributes.Hatch.Name = hatchName;
                                if (label == "HatchOffsetX") circle.Attributes.Hatch.OffsetX = double.Parse(hatchOffsetX);
                                if (label == "HatchOffsetY") circle.Attributes.Hatch.OffsetY = double.Parse(hatchOffsetY);
                                if (label == "HatchScaleX") circle.Attributes.Hatch.ScaleX = double.Parse(hatchScaleX);
                                if (label == "HatchScaleY") circle.Attributes.Hatch.ScaleY = double.Parse(hatchScaleY);
                                circle.Modify();
                                drawing.CommitChanges();
                            }
                        }
                    }
                    catch { }
                }
            }

            public class Polygon : Closed
            {
                private string bulge;

                [Category("Line")]
                public string Bulge
                {
                    get { return bulge; }
                    set { bulge = value; }
                }

                /// <summary>
                /// Get Polygon
                /// </summary>
                public void GetProperties()
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    if (drawingObjectEnum.GetSize() == 1)
                    {
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Polygon)
                            {
                                Tekla.Structures.Drawing.Polygon polygon = (Tekla.Structures.Drawing.Polygon)drawingObjectEnum.Current;
                                bulge = polygon.Bulge.ToString();
                                lineColour = polygon.Attributes.Line.Color;

                                if (polygon.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                                else if (polygon.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                                else if (polygon.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                                else if (polygon.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                                else if (polygon.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                                else if (polygon.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                                else if (polygon.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                                else if (polygon.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                                hatchName = polygon.Attributes.Hatch.Name;
                                hatchAngle = polygon.Attributes.Hatch.Angle.ToString();
                                hatchBackgroundColour = polygon.Attributes.Hatch.BackgroundColor;
                                hatchColour = polygon.Attributes.Hatch.Color;
                                if (polygon.Attributes.Hatch.DrawBackgroundColor) hatchDrawBackgroundColour = Bool.True;
                                else hatchDrawBackgroundColour = Bool.False;

                                if (polygon.Attributes.Hatch.FactorType == 1) hatchFactorType = Bool.True; 
                                else hatchFactorType = Bool.False;
                                hatchOffsetX = polygon.Attributes.Hatch.OffsetX.ToString();
                                hatchOffsetY = polygon.Attributes.Hatch.OffsetY.ToString();
                                hatchScaleX = polygon.Attributes.Hatch.ScaleX.ToString();
                                hatchScaleY = polygon.Attributes.Hatch.ScaleY.ToString();
                            }
                        }
                    }
                    if (drawingObjectEnum.GetSize() > 1)
                    {
                        bulge = "";
                        lineColour = new DrawingColors();
                        lineType = new LineTypeEnum();
                        hatchName = "";
                        hatchAngle = "";
                        hatchBackgroundColour = new DrawingColors();
                        hatchColour = new DrawingColors();
                        hatchDrawBackgroundColour = new Bool();
                        hatchFactorType = new Bool();
                        hatchOffsetX = "";
                        hatchOffsetY = "";
                        hatchScaleX = "";
                        hatchScaleY = "";
                    }
                }

                /// <summary>
                /// Modify Polygon
                /// </summary>
                /// <param name="e"></param>
                public void Modify(PropertyValueChangedEventArgs e)
                {
                    try
                    {
                        string label = e.ChangedItem.Label;
                        Drawing drawing = drawingHandler.GetActiveDrawing();
                        DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                        while (drawingObjectEnum.MoveNext())
                        {
                            if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Polygon)
                            {
                                Tekla.Structures.Drawing.Polygon polygon = (Tekla.Structures.Drawing.Polygon)drawingObjectEnum.Current;

                                if (label == "Bulge") polygon.Bulge = double.Parse(bulge);
                                if (label == "LineColour") polygon.Attributes.Line.Color = lineColour;
                                if (label == "LineType")
                                {
                                    if (lineType == LineTypeEnum.DashDot) polygon.Attributes.Line.Type = LineTypes.DashDot;
                                    else if (lineType == LineTypeEnum.DashDoubleDot) polygon.Attributes.Line.Type = LineTypes.DashDoubleDot;
                                    else if (lineType == LineTypeEnum.DashedLine) polygon.Attributes.Line.Type = LineTypes.DashedLine;
                                    else if (lineType == LineTypeEnum.DottedLine) polygon.Attributes.Line.Type = LineTypes.DottedLine;
                                    else if (lineType == LineTypeEnum.SlashDash) polygon.Attributes.Line.Type = LineTypes.SlashDash;
                                    else if (lineType == LineTypeEnum.SlashedLine) polygon.Attributes.Line.Type = LineTypes.SlashedLine;
                                    else if (lineType == LineTypeEnum.SolidLine) polygon.Attributes.Line.Type = LineTypes.SolidLine;
                                    else if (lineType == LineTypeEnum.UndefinedLine) polygon.Attributes.Line.Type = LineTypes.UndefinedLine;
                                }

                                if (label == "HatchAngle") polygon.Attributes.Hatch.Angle = double.Parse(hatchAngle);
                                if (label == "HatchBackgroundColour") polygon.Attributes.Hatch.BackgroundColor = hatchBackgroundColour;
                                if (label == "HatchColour") polygon.Attributes.Hatch.Color = hatchColour;
                                if (label == "HatchDrawBackgroundColour") 
                                    polygon.Attributes.Hatch.DrawBackgroundColor = bool.Parse(hatchDrawBackgroundColour.ToString());
                                if (label == "HatchFactorType")
                                {
                                    if (hatchFactorType == Bool.True) polygon.Attributes.Hatch.FactorType = 1;
                                    else polygon.Attributes.Hatch.FactorType = 0;
                                }
                                if (label == "HatchName") polygon.Attributes.Hatch.Name = hatchName;
                                if (label == "HatchOffsetX") polygon.Attributes.Hatch.OffsetX = double.Parse(hatchOffsetX);
                                if (label == "HatchOffsetY") polygon.Attributes.Hatch.OffsetY = double.Parse(hatchOffsetY);
                                if (label == "HatchScaleX") polygon.Attributes.Hatch.ScaleX = double.Parse(hatchScaleX);
                                if (label == "HatchScaleY") polygon.Attributes.Hatch.ScaleY = double.Parse(hatchScaleY);
                                polygon.Modify();
                                drawing.CommitChanges();
                            }
                        }
                    }
                    catch { }
                }
            }
        }
    }
}