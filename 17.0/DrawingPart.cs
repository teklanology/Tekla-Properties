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
    class DrawingPart
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private Bool drawCenterLine;
        private Bool drawChamfers;
        //private TrueFalse drawConnectingSideMarks;
        private Bool drawHiddenLines;
        private Bool drawOrientationMark;
        private Bool drawOwnHiddenLines;
        private Bool drawPopMarks;
        private Bool drawReferenceLine;
        private LineTypeEnum visibleLinesType;
        private DrawingColors visibleLinesColour;
        private LineTypeEnum hiddenLinesType;
        private DrawingColors hiddenLinesColour;
        private LineTypeEnum referenceLinesType;
        private DrawingColors referenceLinesColour;

        [Category("Lines")]
        public Bool DrawCenterLine
        {
            get { return drawCenterLine; }
            set { drawCenterLine = value; }
        }

        [Category("Additional Marks")]
        public Bool DrawChamfers
        {
            get { return drawChamfers; }
            set { drawChamfers = value; }
        }
        
        //[Category("Additional Marks")]
        //public bool DrawConnectingSideMarks
        //{
        //    get { return drawConnectingSideMarks; }
        //    set { drawConnectingSideMarks = value; }
        //}

        [Category("Lines")]
        public Bool DrawHiddenLines
        {
            get { return drawHiddenLines; }
            set { drawHiddenLines = value; }
        }

        [Category("Additional Marks")]
        public Bool DrawOrientationMark
        {
            get { return drawOrientationMark; }
            set { drawOrientationMark = value; }
        }

        [Category("Lines")]
        public Bool DrawOwnHiddenLines
        {
            get { return drawOwnHiddenLines; }
            set { drawOwnHiddenLines = value; }
        }

        [Category("Additional Marks")]
        public Bool DrawPopMarks
        {
            get { return drawPopMarks; }
            set { drawPopMarks = value; }
        }

        [Category("Lines")]
        public Bool DrawReferenceLine
        {
            get { return drawReferenceLine; }
            set { drawReferenceLine = value; }
        }

        [Category("Visible Lines")]
        public LineTypeEnum VisibleLinesType
        {
            get { return visibleLinesType; }
            set { visibleLinesType = value; }
        }

        [Category("Visible Lines")]
        public DrawingColors VisibleLinesColour
        {
            get { return visibleLinesColour; }
            set { visibleLinesColour = value; }
        }

        [Category("Hidden Lines")]
        public LineTypeEnum HiddenLinesType
        {
            get { return hiddenLinesType; }
            set { hiddenLinesType = value; }
        }

        [Category("Hidden Lines")]
        public DrawingColors HiddenLinesColour
        {
            get { return hiddenLinesColour; }
            set { hiddenLinesColour = value; }
        }

        [Category("Reference Lines")]
        public LineTypeEnum ReferenceLinesType
        {
            get { return referenceLinesType; }
            set { referenceLinesType = value; }
        }

        [Category("Reference Lines")]
        public DrawingColors ReferenceLinesColour
        {
            get { return referenceLinesColour; }
            set { referenceLinesColour = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Part)
                    {
                        Tekla.Structures.Drawing.Part drawingPart = (Tekla.Structures.Drawing.Part)drawingObjectEnum.Current;

                        if (drawingPart.Attributes.DrawCenterLine) drawCenterLine = Bool.True; else drawCenterLine = Bool.False;
                        if (drawingPart.Attributes.DrawChamfers) drawChamfers = Bool.True; else drawChamfers = Bool.False;
                        //drawConnectingSideMarks = drawingPart.Attributes.DrawConnectingSideMarks;
                        if (drawingPart.Attributes.DrawHiddenLines) drawHiddenLines = Bool.True; else drawHiddenLines = Bool.False;
                        if (drawingPart.Attributes.DrawOrientationMark) drawOrientationMark = Bool.True; else drawOrientationMark = Bool.False;
                        if (drawingPart.Attributes.DrawOwnHiddenLines) drawOwnHiddenLines = Bool.True; else drawOwnHiddenLines = Bool.False;
                        if (drawingPart.Attributes.DrawPopMarks) drawPopMarks = Bool.True; else drawPopMarks = Bool.False;
                        if (drawingPart.Attributes.DrawReferenceLine) drawReferenceLine = Bool.True; else drawReferenceLine = Bool.False;
                        visibleLinesColour = drawingPart.Attributes.VisibleLines.Color;
                        if (drawingPart.Attributes.VisibleLines.Type == LineTypes.DashDot) visibleLinesType = LineTypeEnum.DashDot;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.DashDoubleDot) visibleLinesType = LineTypeEnum.DashDoubleDot;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.DashedLine) visibleLinesType = LineTypeEnum.DashedLine;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.DottedLine) visibleLinesType = LineTypeEnum.DottedLine;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.SlashDash) visibleLinesType = LineTypeEnum.SlashDash;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.SlashedLine) visibleLinesType = LineTypeEnum.SlashedLine;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.SolidLine) visibleLinesType = LineTypeEnum.SolidLine;
                        else if (drawingPart.Attributes.VisibleLines.Type == LineTypes.UndefinedLine) visibleLinesType = LineTypeEnum.UndefinedLine;
                        hiddenLinesColour = drawingPart.Attributes.HiddenLines.Color;
                        if (drawingPart.Attributes.HiddenLines.Type == LineTypes.DashDot) hiddenLinesType = LineTypeEnum.DashDot;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.DashDoubleDot) hiddenLinesType = LineTypeEnum.DashDoubleDot;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.DashedLine) hiddenLinesType = LineTypeEnum.DashedLine;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.DottedLine) hiddenLinesType = LineTypeEnum.DottedLine;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.SlashDash) hiddenLinesType = LineTypeEnum.SlashDash;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.SlashedLine) hiddenLinesType = LineTypeEnum.SlashedLine;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.SolidLine) hiddenLinesType = LineTypeEnum.SolidLine;
                        else if (drawingPart.Attributes.HiddenLines.Type == LineTypes.UndefinedLine) hiddenLinesType = LineTypeEnum.UndefinedLine;
                        referenceLinesColour = drawingPart.Attributes.ReferenceLine.Color;
                        if (drawingPart.Attributes.VisibleLines.Type == LineTypes.DashDot) visibleLinesType = LineTypeEnum.DashDot;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.DashDoubleDot) referenceLinesType = LineTypeEnum.DashDoubleDot;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.DashedLine) referenceLinesType = LineTypeEnum.DashedLine;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.DottedLine) referenceLinesType = LineTypeEnum.DottedLine;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.SlashDash) referenceLinesType = LineTypeEnum.SlashDash;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.SlashedLine) referenceLinesType = LineTypeEnum.SlashedLine;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.SolidLine) referenceLinesType = LineTypeEnum.SolidLine;
                        else if (drawingPart.Attributes.ReferenceLine.Type == LineTypes.UndefinedLine) referenceLinesType = LineTypeEnum.UndefinedLine;
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                drawCenterLine = new Bool();
                drawChamfers = new Bool();
                drawHiddenLines = new Bool();
                drawOrientationMark = new Bool();
                drawOwnHiddenLines = new Bool();
                DrawPopMarks = new Bool();
                drawReferenceLine = new Bool();
                visibleLinesColour = new DrawingColors();
                visibleLinesType = new LineTypeEnum();
                hiddenLinesColour = new DrawingColors();
                hiddenLinesType = new LineTypeEnum();
                referenceLinesColour = new DrawingColors();
                referenceLinesType = new LineTypeEnum();
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Part)
                {
                    Tekla.Structures.Drawing.Part drawingPart = (Tekla.Structures.Drawing.Part)drawingObjectEnum.Current;
                    if (label == "DrawCenterLine") drawingPart.Attributes.DrawCenterLine = bool.Parse(drawCenterLine.ToString());
                    if (label == "DrawChamfers") drawingPart.Attributes.DrawChamfers = bool.Parse(drawChamfers.ToString());
                    //if (label == "Draw Connecting Side Marks") drawingPart.Attributes.DrawConnectingSideMarks = drawConnectingSideMarks;
                    if (label == "DrawHiddenLines") drawingPart.Attributes.DrawHiddenLines = bool.Parse(drawHiddenLines.ToString());
                    if (label == "DrawOrientationMark") drawingPart.Attributes.DrawOrientationMark = bool.Parse(drawOrientationMark.ToString());
                    if (label == "DrawOwnHiddenLines") drawingPart.Attributes.DrawOwnHiddenLines = bool.Parse(drawOwnHiddenLines.ToString());
                    if (label == "DrawPopMarks") drawingPart.Attributes.DrawPopMarks = bool.Parse(drawPopMarks.ToString());
                    if (label == "DrawReferenceLine") drawingPart.Attributes.DrawReferenceLine = bool.Parse(drawReferenceLine.ToString());
                    if (label == "VisibleLinesColour") drawingPart.Attributes.VisibleLines.Color = visibleLinesColour;
                    if (label == "VisibleLinesType")
                    {
                        if (visibleLinesType == LineTypeEnum.DashDot) drawingPart.Attributes.VisibleLines.Type = LineTypes.DashDot;
                        else if (visibleLinesType == LineTypeEnum.DashDoubleDot) drawingPart.Attributes.VisibleLines.Type = LineTypes.DashDoubleDot;
                        else if (visibleLinesType == LineTypeEnum.DashedLine) drawingPart.Attributes.VisibleLines.Type = LineTypes.DashedLine;
                        else if (visibleLinesType == LineTypeEnum.DottedLine) drawingPart.Attributes.VisibleLines.Type = LineTypes.DottedLine;
                        else if (visibleLinesType == LineTypeEnum.SlashDash) drawingPart.Attributes.VisibleLines.Type = LineTypes.SlashDash;
                        else if (visibleLinesType == LineTypeEnum.SlashedLine) drawingPart.Attributes.VisibleLines.Type = LineTypes.SlashedLine;
                        else if (visibleLinesType == LineTypeEnum.SolidLine) drawingPart.Attributes.VisibleLines.Type = LineTypes.SolidLine;
                        else if (visibleLinesType == LineTypeEnum.UndefinedLine) drawingPart.Attributes.VisibleLines.Type = LineTypes.UndefinedLine;
                    }
                    if (label == "HiddenLinesColour") drawingPart.Attributes.HiddenLines.Color = hiddenLinesColour;
                    if (label == "HiddenLinesType")
                    {
                        if (hiddenLinesType == LineTypeEnum.DashDot) drawingPart.Attributes.HiddenLines.Type = LineTypes.DashDot;
                        else if (hiddenLinesType == LineTypeEnum.DashDoubleDot) drawingPart.Attributes.HiddenLines.Type = LineTypes.DashDoubleDot;
                        else if (hiddenLinesType == LineTypeEnum.DashedLine) drawingPart.Attributes.HiddenLines.Type = LineTypes.DashedLine;
                        else if (hiddenLinesType == LineTypeEnum.DottedLine) drawingPart.Attributes.HiddenLines.Type = LineTypes.DottedLine;
                        else if (hiddenLinesType == LineTypeEnum.SlashDash) drawingPart.Attributes.HiddenLines.Type = LineTypes.SlashDash;
                        else if (hiddenLinesType == LineTypeEnum.SlashedLine) drawingPart.Attributes.HiddenLines.Type = LineTypes.SlashedLine;
                        else if (hiddenLinesType == LineTypeEnum.SolidLine) drawingPart.Attributes.HiddenLines.Type = LineTypes.SolidLine;
                        else if (hiddenLinesType == LineTypeEnum.UndefinedLine) drawingPart.Attributes.HiddenLines.Type = LineTypes.UndefinedLine;
                    }
                    if (label == "ReferenceLinesColour") drawingPart.Attributes.ReferenceLine.Color = referenceLinesColour;
                    if (label == "ReferenceLinesType") 
                    {
                        if (referenceLinesType == LineTypeEnum.DashDot) drawingPart.Attributes.ReferenceLine.Type = LineTypes.DashDot;
                        else if (referenceLinesType == LineTypeEnum.DashDoubleDot) drawingPart.Attributes.ReferenceLine.Type = LineTypes.DashDoubleDot;
                        else if (referenceLinesType == LineTypeEnum.DashedLine) drawingPart.Attributes.ReferenceLine.Type = LineTypes.DashedLine;
                        else if (referenceLinesType == LineTypeEnum.DottedLine) drawingPart.Attributes.ReferenceLine.Type = LineTypes.DottedLine;
                        else if (referenceLinesType == LineTypeEnum.SlashDash) drawingPart.Attributes.ReferenceLine.Type = LineTypes.SlashDash;
                        else if (referenceLinesType == LineTypeEnum.SlashedLine) drawingPart.Attributes.ReferenceLine.Type = LineTypes.SlashedLine;
                        else if (referenceLinesType == LineTypeEnum.SolidLine) drawingPart.Attributes.ReferenceLine.Type = LineTypes.SolidLine;
                        else if (referenceLinesType == LineTypeEnum.UndefinedLine) drawingPart.Attributes.ReferenceLine.Type = LineTypes.UndefinedLine;
                    }
                    drawingPart.Modify();
                    drawing.CommitChanges();
                }
            }
        }
    }
}
