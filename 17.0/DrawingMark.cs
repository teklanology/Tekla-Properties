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
    class DrawingMark
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private DrawingColors frameColour;
        private FrameTypeEnum frameType;

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

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Mark)
                    {
                        Tekla.Structures.Drawing.Mark drawingMark = (Tekla.Structures.Drawing.Mark)drawingObjectEnum.Current;
                        frameColour = drawingMark.Attributes.Frame.Color;
                        frameType = (FrameTypeEnum)Enum.Parse(typeof(FrameTypeEnum), drawingMark.Attributes.Frame.Type.ToString());
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                frameColour = new DrawingColors();
                FrameType = new FrameTypeEnum();
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.Mark)
                {
                    Tekla.Structures.Drawing.Mark drawingMark = (Tekla.Structures.Drawing.Mark)drawingObjectEnum.Current;
                    if (label == "FrameColour") drawingMark.Attributes.Frame.Color = frameColour;
                    if (label == "FrameType") drawingMark.Attributes.Frame.Type = (FrameTypes)Enum.Parse(typeof(FrameTypes), frameType.ToString());

                    drawingMark.Modify();
                    drawing.CommitChanges();
                }
            }
        }

    }
}
