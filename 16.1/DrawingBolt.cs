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
    class DrawingBolt
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private Bolt.Representation boltRepresentation;
        private Bool symbolContainsAxis;
        private Bool symbolContainsHole;
        private DrawingColors colour;

        [Category("Content")]
        public Bolt.Representation BoltRepresentation
        {
            get { return boltRepresentation; }
            set { boltRepresentation = value; }
        }

        [Category("Content")]
        public Bool SymbolContainsAxis
        {
            get { return symbolContainsAxis; }
            set { symbolContainsAxis = value; }
        }

        [Category("Content")]
        public Bool SymbolContainsHole
        {
            get { return symbolContainsHole; }
            set { symbolContainsHole = value; }
        }

        [Category("Appearance")]
        public DrawingColors Colour
        {
            get { return colour; }
            set { colour = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Bolt)
                    {
                        Tekla.Structures.Drawing.Bolt drawingBolt = (Tekla.Structures.Drawing.Bolt)drawingObjectEnum.Current;
                        boltRepresentation = drawingBolt.Attributes.Representation;
                        if (drawingBolt.Attributes.SymbolContainsAxis) symbolContainsAxis = Bool.True; else symbolContainsAxis = Bool.False;
                        if (drawingBolt.Attributes.SymbolContainsHole) symbolContainsHole = Bool.True; else symbolContainsHole = Bool.False;
                        colour = drawingBolt.Attributes.Color;
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                boltRepresentation = new Bolt.Representation();
                symbolContainsAxis = new Bool();
                SymbolContainsHole = new Bool();
                colour = new DrawingColors();
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Bolt)
                {
                    Tekla.Structures.Drawing.Bolt drawingBolt = (Tekla.Structures.Drawing.Bolt)drawingObjectEnum.Current;
                    if (label == "BoltRepresentation") drawingBolt.Attributes.Representation = boltRepresentation;
                    if (label == "SymbolContainsAxis") drawingBolt.Attributes.SymbolContainsAxis = bool.Parse(symbolContainsAxis.ToString());
                    if (label == "SymbolContainsHole") drawingBolt.Attributes.SymbolContainsHole = bool.Parse(symbolContainsHole.ToString());
                    if (label == "Colour") drawingBolt.Attributes.Color = colour;
                    drawingBolt.Modify();
                    drawing.CommitChanges();
                }
            }
        }

    }
}
