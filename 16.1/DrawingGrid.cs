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
    class DrawingGrid
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private Bool drawOnlyTextLabelsNotGridLines;
        private Bool drawTextAtEndOfGridLine;
        private Bool drawTextAtStartOfGridLine;
        private DrawingColors fontColour;
        private string fontHeight;
        private Bool fontBold;
        private Bool fontItalic;
        private string fontName;
        private DrawingColors frameColour;
        private FrameTypeEnum frameType;
        private DrawingColors lineColour;
        private LineTypeEnum lineType;
        private string offsetAtEndOfLine;
        private string offsetAtStartOfLine;
        private string gridLabelText;

        [Category("...")]
        public Bool DrawOnlyTextLabelsNotGridLines
        {
            get { return drawOnlyTextLabelsNotGridLines; }
            set { drawOnlyTextLabelsNotGridLines = value; }
        }

        [Category("...")]
        public Bool DrawTextAtEndOfGridLine
        {
            get { return drawTextAtEndOfGridLine; }
            set { drawTextAtEndOfGridLine = value; }
        }

        [Category("...")]
        public Bool DrawTextAtStartOfGridLine
        {
            get { return drawTextAtStartOfGridLine; }
            set { drawTextAtStartOfGridLine = value; }
        }

        [Category("Text")]
        public DrawingColors FontColour
        {
            get { return fontColour; }
            set { fontColour = value; }
        }

        [Category("Text")]
        public string FontHeight
        {
            get { return fontHeight; }
            set { fontHeight = value; }
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

        [Category("Text")]
        public string FontName
        {
            get { return fontName; }
            set { fontName = value; }
        }

        [Category("Text")]
        public DrawingColors FrameColour
        {
            get { return frameColour; }
            set { frameColour = value; }
        }

        [Category("Text")]
        public FrameTypeEnum FrameType
        {
            get { return frameType; }
            set { frameType = value; }
        }

        [Category("Grid Line")]
        public DrawingColors LineColour
        {
            get { return lineColour; }
            set { lineColour = value; }
        }

        [Category("Grid Line")]
        public LineTypeEnum LineType
        {
            get { return lineType; }
            set { lineType = value; }
        }

        public string OffsetAtEndOfLine
        {
            get { return offsetAtEndOfLine; }
            set { offsetAtEndOfLine = value; }
        }

        [Category("...")]
        public string OffsetAtStartOfLine
        {
            get { return offsetAtStartOfLine; }
            set { offsetAtStartOfLine = value; }
        }

        [Category("...")]
        public string GridLabelText
        {
            get { return gridLabelText; }
            set { gridLabelText = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.GridLine)
                    {
                        Tekla.Structures.Drawing.GridLine grid = (Tekla.Structures.Drawing.GridLine)drawingObjectEnum.Current;
                        if (grid.Attributes.DrawOnlyTextLabelsNotGridLines) drawOnlyTextLabelsNotGridLines = Bool.True;
                        else drawOnlyTextLabelsNotGridLines = Bool.False;

                        if (grid.Attributes.DrawTextAtEndOfGridLine) drawTextAtEndOfGridLine = Bool.True;
                        else drawTextAtEndOfGridLine = Bool.False;

                        if (grid.Attributes.DrawTextAtStartOfGridLine) drawTextAtStartOfGridLine = Bool.True;
                        else drawTextAtStartOfGridLine = Bool.False;

                        if (grid.Attributes.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;

                        fontColour = grid.Attributes.Font.Color;
                        fontHeight = grid.Attributes.Font.Height.ToString();

                        if (grid.Attributes.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;

                        fontName = grid.Attributes.Font.Name;
                        frameColour = grid.Attributes.Frame.Color;
                        if (grid.Attributes.Frame.Type == FrameTypes.Circle) frameType = FrameTypeEnum.Circle;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Diamond) frameType = FrameTypeEnum.Diamond;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Hexagon) frameType = FrameTypeEnum.Hexagon;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Line) frameType = FrameTypeEnum.Line;
                        else if (grid.Attributes.Frame.Type == FrameTypes.None) frameType = FrameTypeEnum.None;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Rectangular) frameType = FrameTypeEnum.Rectangular;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Round) frameType = FrameTypeEnum.Round;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Sharpened) frameType = FrameTypeEnum.Sharpened;
                        else if (grid.Attributes.Frame.Type == FrameTypes.Triangle) frameType = FrameTypeEnum.Triangle;
                        lineColour = grid.Attributes.Line.Color;

                        if (grid.Attributes.Line.Type == LineTypes.DashDot) lineType = LineTypeEnum.DashDot;
                        else if (grid.Attributes.Line.Type == LineTypes.DashDoubleDot) lineType = LineTypeEnum.DashDoubleDot;
                        else if (grid.Attributes.Line.Type == LineTypes.DashedLine) lineType = LineTypeEnum.DashedLine;
                        else if (grid.Attributes.Line.Type == LineTypes.DottedLine) lineType = LineTypeEnum.DottedLine;
                        else if (grid.Attributes.Line.Type == LineTypes.SlashDash) lineType = LineTypeEnum.SlashDash;
                        else if (grid.Attributes.Line.Type == LineTypes.SlashedLine) lineType = LineTypeEnum.SlashedLine;
                        else if (grid.Attributes.Line.Type == LineTypes.SolidLine) lineType = LineTypeEnum.SolidLine;
                        else if (grid.Attributes.Line.Type == LineTypes.UndefinedLine) lineType = LineTypeEnum.UndefinedLine;

                        offsetAtEndOfLine = grid.Attributes.OffsetAtEndOfLine.ToString();
                        offsetAtStartOfLine = grid.Attributes.OffsetAtStartOfLine.ToString();
                        gridLabelText = grid.EndLabel.GridLabelText;
                    }
                }
            }
            if (drawingObjectEnum.GetSize() > 1)
            {
                drawOnlyTextLabelsNotGridLines = new Bool();
                drawTextAtEndOfGridLine = new Bool();
                drawTextAtStartOfGridLine = new Bool();
                fontColour = new DrawingColors();
                fontHeight = "";
                fontBold = new Bool();
                fontItalic = new Bool();
                fontName = "";
                frameColour = new DrawingColors();
                frameType = new FrameTypeEnum();
                lineColour = new DrawingColors();
                lineType = new LineTypeEnum();
                offsetAtEndOfLine = "";
                offsetAtStartOfLine = "";
                gridLabelText = "";
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.GridLine)
                {
                    Tekla.Structures.Drawing.GridLine grid = (Tekla.Structures.Drawing.GridLine)drawingObjectEnum.Current;
                    if (label == "DrawOnlyTextLabelsNotGridLines") 
                        grid.Attributes.DrawOnlyTextLabelsNotGridLines = bool.Parse(DrawOnlyTextLabelsNotGridLines.ToString());
                    if (label == "DrawTextAtEndOfGridLine")
                        grid.Attributes.DrawTextAtEndOfGridLine = bool.Parse(drawTextAtEndOfGridLine.ToString());
                    if (label == "DrawTextAtStartOfGridLine")
                        grid.Attributes.DrawTextAtStartOfGridLine = bool.Parse(drawTextAtStartOfGridLine.ToString());
                    if (label == "FontBold") grid.Attributes.Font.Bold = bool.Parse(fontBold.ToString());
                    if (label == "FontColour") grid.Attributes.Font.Color = fontColour;
                    if (label == "FontHeight") grid.Attributes.Font.Height = double.Parse(fontHeight);
                    if (label == "FontItalic") grid.Attributes.Font.Italic = bool.Parse(fontItalic.ToString());
                    if (label == "FontName") grid.Attributes.Font.Name = fontName;
                    if (label == "FrameColour") grid.Attributes.Frame.Color = frameColour;
                    if (label == "FrameType")
                    {
                        if (frameType == FrameTypeEnum.Circle) grid.Attributes.Frame.Type = FrameTypes.Circle;
                        else if (frameType == FrameTypeEnum.Diamond) grid.Attributes.Frame.Type = FrameTypes.Diamond;
                        else if (frameType == FrameTypeEnum.Hexagon) grid.Attributes.Frame.Type = FrameTypes.Hexagon;
                        else if (frameType == FrameTypeEnum.Line) grid.Attributes.Frame.Type = FrameTypes.Line;
                        else if (frameType == FrameTypeEnum.None) grid.Attributes.Frame.Type = FrameTypes.None;
                        else if (frameType == FrameTypeEnum.Rectangular) grid.Attributes.Frame.Type = FrameTypes.Rectangular;
                        else if (frameType == FrameTypeEnum.Round) grid.Attributes.Frame.Type = FrameTypes.Round;
                        else if (frameType == FrameTypeEnum.Sharpened) grid.Attributes.Frame.Type = FrameTypes.Sharpened;
                        else if (frameType == FrameTypeEnum.Triangle) grid.Attributes.Frame.Type = FrameTypes.Triangle;
                    }

                    if (label == "LineColour") grid.Attributes.Line.Color = lineColour;
                    if (label == "LineType")
                    {
                        if (lineType == LineTypeEnum.DashDot) grid.Attributes.Line.Type = LineTypes.DashDot;
                        else if (lineType == LineTypeEnum.DashDoubleDot) grid.Attributes.Line.Type = LineTypes.DashDoubleDot;
                        else if (lineType == LineTypeEnum.DashedLine) grid.Attributes.Line.Type = LineTypes.DashedLine;
                        else if (lineType == LineTypeEnum.DottedLine) grid.Attributes.Line.Type = LineTypes.DottedLine;
                        else if (lineType == LineTypeEnum.SlashDash) grid.Attributes.Line.Type = LineTypes.SlashDash;
                        else if (lineType == LineTypeEnum.SlashedLine) grid.Attributes.Line.Type = LineTypes.SlashedLine;
                        else if (lineType == LineTypeEnum.SolidLine) grid.Attributes.Line.Type = LineTypes.SolidLine;
                        else if (lineType == LineTypeEnum.UndefinedLine) grid.Attributes.Line.Type = LineTypes.UndefinedLine;
                    }
                    if (label == "OffsetAtEndOfLine") grid.Attributes.OffsetAtEndOfLine = double.Parse(offsetAtEndOfLine);
                    if (label == "OffsetAtStartOfLine") grid.Attributes.OffsetAtStartOfLine = double.Parse(offsetAtStartOfLine);
                    grid.Modify();
                    drawing.CommitChanges();
                }
            }
        }

    }
}
