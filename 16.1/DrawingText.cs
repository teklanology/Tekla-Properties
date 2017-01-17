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
    public class DrawingText
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        
        private PreferredPlacingEnum preferredPlacingType;
        private string textString;
        private string textAngle;
        private DrawingColors fontColour;
        private string fontHeight;
        private string fontName;
        private Bool fontBold;
        private Bool fontItalic;
        private TextAlignmentEnum textAlignment;
        private FrameTypeEnum frameType;
        private DrawingColors frameColour;
        private ArrowheadTypeEnum arrowheadType;
        private ArrowheadPositionEnum arrowheadPosition;
        private string arrowheadWidth;
        private string arrowheadHeight;

        [Category("Text")]
        public string TextString
        {
            get { return textString; }
            set { textString = value; }
        }

        [Category("Text")]
        public string TextAngle
        {
            get { return textAngle; }
            set { textAngle = value; }
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
        public string FontName
        {
            get { return fontName; }
            set { fontName = value; }
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
        public TextAlignmentEnum TextAlignment
        {
            get { return textAlignment; }
            set { textAlignment = value; }
        }

        [Category("Frame")]
        public FrameTypeEnum FrameType
        {
            get { return frameType; }
            set { frameType = value; }
        }

        [Category("Frame")]
        public DrawingColors FrameColour
        {
            get { return frameColour; }
            set { frameColour = value; }
        }

        [Category("Arrowhead")]
        public ArrowheadTypeEnum ArrowheadType
        {
            get { return arrowheadType; }
            set { arrowheadType = value; }
        }

        [Category("Arrowhead")]
        public ArrowheadPositionEnum ArrowheadPosition
        {
            get { return arrowheadPosition; }
            set { arrowheadPosition = value; }
        }

        [Category("Arrowhead")]
        public string ArrowheadWidth
        {
            get { return arrowheadWidth; }
            set { arrowheadWidth = value; }
        }

        [Category("Arrowhead")]
        public string ArrowheadHeight
        {
            get { return arrowheadHeight; }
            set { arrowheadHeight = value; }
        }

        [Category("Placing")]
        public PreferredPlacingEnum PreferredPlacingType
        {
            get { return preferredPlacingType; }
            set { preferredPlacingType = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Text)
                    {
                        Tekla.Structures.Drawing.Text text = (Tekla.Structures.Drawing.Text)drawingObjectEnum.Current;

                        Text.TextAttributes textAttributes = text.Attributes;
                        if (textAttributes.PreferredPlacing.ToString() == "Tekla.Structures.Drawing.PointPlacingType")
                            preferredPlacingType = PreferredPlacingEnum.PointPlacingType;
                        else if (textAttributes.PreferredPlacing.ToString() == "Tekla.Structures.Drawing.LeaderLinePlacingType")
                            preferredPlacingType = PreferredPlacingEnum.LeaderLinePlacingType;

                        textString = text.TextString.Replace("\n", "¬");

                        fontColour = text.Attributes.Font.Color;
                        fontHeight = text.Attributes.Font.Height.ToString();
                        fontName = text.Attributes.Font.Name;
                        if (text.Attributes.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;
                        if (text.Attributes.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;
                        textAngle = text.Attributes.Angle.ToString();

                        if (text.Attributes.Alignment == Tekla.Structures.Drawing.TextAlignment.Center) textAlignment = TextAlignmentEnum.Center;
                        else if (text.Attributes.Alignment == Tekla.Structures.Drawing.TextAlignment.Left) textAlignment = TextAlignmentEnum.Left;
                        else if (text.Attributes.Alignment == Tekla.Structures.Drawing.TextAlignment.Right) textAlignment = TextAlignmentEnum.Right;

                        if (text.Attributes.Frame.Type == FrameTypes.Circle) frameType = FrameTypeEnum.Circle;
                        else if (text.Attributes.Frame.Type == FrameTypes.Diamond) frameType = FrameTypeEnum.Diamond;
                        else if (text.Attributes.Frame.Type == FrameTypes.Hexagon) frameType = FrameTypeEnum.Hexagon;
                        else if (text.Attributes.Frame.Type == FrameTypes.Line) frameType = FrameTypeEnum.Line;
                        else if (text.Attributes.Frame.Type == FrameTypes.None) frameType = FrameTypeEnum.None;
                        else if (text.Attributes.Frame.Type == FrameTypes.Rectangular) frameType = FrameTypeEnum.Rectangular;
                        else if (text.Attributes.Frame.Type == FrameTypes.Round) frameType = FrameTypeEnum.Round;
                        else if (text.Attributes.Frame.Type == FrameTypes.Sharpened) frameType = FrameTypeEnum.Sharpened;
                        else if (text.Attributes.Frame.Type == FrameTypes.Triangle) frameType = FrameTypeEnum.Triangle;

                        frameColour = text.Attributes.Frame.Color;

                        if (text.Attributes.ArrowHead.Head == ArrowheadTypes.CircleArrow) arrowheadType = ArrowheadTypeEnum.CircleArrow;
                        else if (text.Attributes.ArrowHead.Head == ArrowheadTypes.FilledArrow) arrowheadType = ArrowheadTypeEnum.FilledArrow;
                        else if (text.Attributes.ArrowHead.Head == ArrowheadTypes.FilledCircleArrow) arrowheadType = ArrowheadTypeEnum.FilledCircleArrow;
                        else if (text.Attributes.ArrowHead.Head == ArrowheadTypes.LineArrow) arrowheadType = ArrowheadTypeEnum.LineArrow;

                        if (text.Attributes.ArrowHead.ArrowPosition == ArrowheadPositions.Both) arrowheadPosition = ArrowheadPositionEnum.Both;
                        else if (text.Attributes.ArrowHead.ArrowPosition == ArrowheadPositions.End) arrowheadPosition = ArrowheadPositionEnum.End;
                        else if (text.Attributes.ArrowHead.ArrowPosition == ArrowheadPositions.None) arrowheadPosition = ArrowheadPositionEnum.None;
                        else if (text.Attributes.ArrowHead.ArrowPosition == ArrowheadPositions.Start) arrowheadPosition = ArrowheadPositionEnum.Start;

                        arrowheadWidth = text.Attributes.ArrowHead.Width.ToString();
                        arrowheadHeight = text.Attributes.ArrowHead.Height.ToString();
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                preferredPlacingType = new PreferredPlacingEnum();
                textString = "";
                FontColour = new DrawingColors();
                fontHeight = "";
                fontName = "";
                fontBold = new Bool();
                fontItalic = new Bool();
                textAngle = "";
                textAlignment = new TextAlignmentEnum();
                frameType = new FrameTypeEnum();
                FrameColour = new DrawingColors();
                arrowheadType = new ArrowheadTypeEnum();
                arrowheadPosition = new ArrowheadPositionEnum();
                arrowheadWidth = "";
                arrowheadHeight = "";
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            try
            {
                string label = e.ChangedItem.Label;
                Drawing drawing = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Text)
                    {
                        Tekla.Structures.Drawing.Text text = (Tekla.Structures.Drawing.Text)drawingObjectEnum.Current;

                        if (label == "PreferredPlacingType")
                        {
                            Text.TextAttributes textAttributes = text.Attributes;
                            if (preferredPlacingType == PreferredPlacingEnum.PointPlacingType)
                                textAttributes.PreferredPlacing = PreferredPlacingTypes.PointPlacingType();
                            else if (preferredPlacingType == PreferredPlacingEnum.LeaderLinePlacingType)
                                textAttributes.PreferredPlacing = PreferredPlacingTypes.LeaderLinePlacingType();
                        }

                        if (label == "TextString")
                        {
                            string formattedText = "";
                            string[] split = textString.Split(new char[] { '¬' });
                            foreach (string temp in split)
                                formattedText = formattedText + temp + Environment.NewLine;

                            formattedText = formattedText.Trim();
                            text.TextString = formattedText;
                        }

                        if (label == "FontColour") text.Attributes.Font.Color = fontColour;
                        if (label == "FontHeight") text.Attributes.Font.Height = double.Parse(fontHeight);
                        if (label == "FontName") text.Attributes.Font.Name = fontName;
                        if (label == "FontBold") text.Attributes.Font.Bold = bool.Parse(fontBold.ToString());
                        if (label == "FontItalic") text.Attributes.Font.Italic = bool.Parse(fontItalic.ToString());
                        if (label == "TextAngle") text.Attributes.Angle = double.Parse(textAngle);
                        if (label == "TextAlignment")
                        {
                            if (textAlignment == TextAlignmentEnum.Center) text.Attributes.Alignment = Tekla.Structures.Drawing.TextAlignment.Center;
                            else if (textAlignment == TextAlignmentEnum.Left) text.Attributes.Alignment = Tekla.Structures.Drawing.TextAlignment.Left;
                            else if (textAlignment == TextAlignmentEnum.Right) text.Attributes.Alignment = Tekla.Structures.Drawing.TextAlignment.Right;
                        }

                        if (label == "FrameType")
                        {
                            if (frameType == FrameTypeEnum.Circle) text.Attributes.Frame.Type = FrameTypes.Circle;
                            else if (frameType == FrameTypeEnum.Diamond) text.Attributes.Frame.Type = FrameTypes.Diamond;
                            else if (frameType == FrameTypeEnum.Hexagon) text.Attributes.Frame.Type = FrameTypes.Hexagon;
                            else if (frameType == FrameTypeEnum.Line) text.Attributes.Frame.Type = FrameTypes.Line;
                            else if (frameType == FrameTypeEnum.None) text.Attributes.Frame.Type = FrameTypes.None;
                            else if (frameType == FrameTypeEnum.Rectangular) text.Attributes.Frame.Type = FrameTypes.Rectangular;
                            else if (frameType == FrameTypeEnum.Round) text.Attributes.Frame.Type = FrameTypes.Round;
                            else if (frameType == FrameTypeEnum.Sharpened) text.Attributes.Frame.Type = FrameTypes.Sharpened;
                            else if (frameType == FrameTypeEnum.Triangle) text.Attributes.Frame.Type = FrameTypes.Triangle;
                        }

                        if (label == "FrameColour") text.Attributes.Frame.Color = frameColour;
                        if (label == "ArrowheadType")
                        {
                            if (arrowheadType == ArrowheadTypeEnum.CircleArrow) text.Attributes.ArrowHead.Head = ArrowheadTypes.CircleArrow;
                            else if (arrowheadType == ArrowheadTypeEnum.FilledArrow) text.Attributes.ArrowHead.Head = ArrowheadTypes.FilledArrow;
                            else if (arrowheadType == ArrowheadTypeEnum.FilledCircleArrow) text.Attributes.ArrowHead.Head = ArrowheadTypes.FilledCircleArrow;
                            else if (arrowheadType == ArrowheadTypeEnum.LineArrow) text.Attributes.ArrowHead.Head = ArrowheadTypes.LineArrow;
                        }

                        if (label == "ArrowheadPosition")
                        {
                            if (arrowheadPosition == ArrowheadPositionEnum.Both) text.Attributes.ArrowHead.ArrowPosition = ArrowheadPositions.Both;
                            else if (arrowheadPosition == ArrowheadPositionEnum.End) text.Attributes.ArrowHead.ArrowPosition = ArrowheadPositions.End;
                            else if (arrowheadPosition == ArrowheadPositionEnum.None) text.Attributes.ArrowHead.ArrowPosition = ArrowheadPositions.None;
                            else if (arrowheadPosition == ArrowheadPositionEnum.Start) text.Attributes.ArrowHead.ArrowPosition = ArrowheadPositions.Start;
                        }

                        if (label == "ArrowheadWidth") text.Attributes.ArrowHead.Width = double.Parse(arrowheadWidth);
                        if (label == "ArrowheadHeight") text.Attributes.ArrowHead.Height = double.Parse(arrowheadHeight);
                        text.Modify();
                        drawing.CommitChanges();
                    }
                }
            }
            catch { }
        }
    }
}
