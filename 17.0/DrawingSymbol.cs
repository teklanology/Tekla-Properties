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
    class DrawingSymbol
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private string angle;
        private DrawingColors colour;
        private DrawingColors frameColour;
        private FrameTypeEnum frameType;
        private string height;
        private Tekla.Structures.Geometry3d.Point insertionPoint;
        private PlacingTypeEnum placingType;
        private string symbolFile;
        private string symbolIndex;

        [Category("Symbol")]
        public string Angle
        {
            get { return angle; }
            set { angle = value; }
        }

        [Category("Symbol")]
        public DrawingColors Colour
        {
            get { return colour; }
            set { colour = value; }
        }

        [Category("Frame")]
        public DrawingColors FrameColour
        {
            get { return frameColour; }
            set { frameColour = value; }
        }

        [Category("Frame")]
        public FrameTypeEnum FrameType
        {
            get { return frameType; }
            set { frameType = value; }
        }

        [Category("Placing")]
        public Tekla.Structures.Geometry3d.Point InsertionPoint
        {
            get { return insertionPoint; }
            set { insertionPoint = value; }
        }

        [Category("Symbol")]
        public string SymbolHeight
        {
            get { return height;}
            set { height = value; }

        }

        [Category("Placing")]
        public PlacingTypeEnum PlacingType
        {
            get { return placingType; }
            set { placingType = value; }
        }

        [Category("Symbol Selection")]
        public string SymbolFile
        {
            get { return symbolFile; }
            set { symbolFile = value; }
        }

        [Category("Symbol Selection")]
        public string SymbolIndex
        {
            get { return symbolIndex; }
            set { symbolIndex = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Symbol)
                    {
                        Tekla.Structures.Drawing.Symbol drawingSymbol = (Tekla.Structures.Drawing.Symbol)drawingObjectEnum.Current;
                        angle = drawingSymbol.Attributes.Angle.ToString();
                        colour = drawingSymbol.Attributes.Color;
                        frameColour = drawingSymbol.Attributes.Frame.Color;
                        frameType = (FrameTypeEnum)Enum.Parse(typeof(FrameTypeEnum), drawingSymbol.Attributes.Frame.Type.ToString());
                        height = drawingSymbol.Attributes.Height.ToString();
                        insertionPoint = drawingSymbol.InsertionPoint;

                        string placing = drawingSymbol.Placing.ToString().Replace("Tekla.Structures.Drawing.", "");
                        placingType = (PlacingTypeEnum)Enum.Parse(typeof(PlacingTypeEnum), placing);

                        symbolFile = drawingSymbol.SymbolInfo.SymbolFile;
                        symbolIndex = drawingSymbol.SymbolInfo.SymbolIndex.ToString();
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                angle = "";
                colour = new DrawingColors();
                frameColour = new DrawingColors();
                frameType = new FrameTypeEnum();
                height = "";
                insertionPoint = null;
                placingType = new PlacingTypeEnum();
                symbolFile = "";
                symbolIndex = "";
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Symbol)
                {
                    Tekla.Structures.Drawing.Symbol drawingSymbol = (Tekla.Structures.Drawing.Symbol)drawingObjectEnum.Current;
                    if (label == "Angle") drawingSymbol.Attributes.Angle = double.Parse(angle);
                    if (label == "Colour") drawingSymbol.Attributes.Color = colour;
                    if (label == "FrameColour") drawingSymbol.Attributes.Frame.Color = frameColour;
                    if (label == "FrameType") drawingSymbol.Attributes.Frame.Type = (FrameTypes)Enum.Parse(typeof(FrameTypes), frameType.ToString());
                    if (label == "SymbolHeight") drawingSymbol.Attributes.Height = double.Parse(height);

                    if (label == "PlacingType")
                    {
                        if (placingType == PlacingTypeEnum.PointPlacing) drawingSymbol.Placing = PlacingTypes.PointPlacing();
                        else if (placingType == PlacingTypeEnum.LeaderLinePlacing) drawingSymbol.Placing = PlacingTypes.LeaderLinePlacing(drawingSymbol.InsertionPoint);
                    }

                    if (label == "SymbolFile") drawingSymbol.SymbolInfo.SymbolFile = symbolFile;
                    if (label == "SymbolIndex") drawingSymbol.SymbolInfo.SymbolIndex = int.Parse(symbolIndex);
                    drawingSymbol.Modify();
                    drawing.CommitChanges();
                }
            }
        }
    }
}
