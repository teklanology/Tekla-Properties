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
    class ActiveDrawing
    {
        #region properties

        private string name;
        private string title1;
        private string title2;
        private string title3;
        private string height;
        private string width;
        private string userfield1;
        private string userfield2;
        private string userfield3;
        private string userfield4;
        private string userfield5;
        private string userfield6;
        private string userfield7;
        private string userfield8;
        private string drawnBy;
        private string checkedBy;

        [Category("Drawing Properties")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [Category("Drawing Properties")]
        public string Title1
        {
            get { return title1; }
            set { title1 = value; }
        }

        [Category("Drawing Properties")]
        public string Title2
        {
            get { return title2; }
            set { title2 = value; }
        }

        [Category("Drawing Properties")]
        public string Title3
        {
            get { return title3; }
            set { title3 = value; }
        }

        [Category("Drawing Properties")]
        [ReadOnly(true)]
        public string Height
        {
            get { return height; }
            set { height = value; }
        }

        [Category("Drawing Properties")]
        [ReadOnly(true)]
        public string Width
        {
            get { return width; }
            set { width = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield1
        {
            get { return userfield1; }
            set { userfield1 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield2
        {
            get { return userfield2; }
            set { userfield2 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield3
        {
            get { return userfield3; }
            set { userfield3 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield4
        {
            get { return userfield4; }
            set { userfield4 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield5
        {
            get { return userfield5; }
            set { userfield5 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield6
        {
            get { return userfield6; }
            set { userfield6 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield7
        {
            get { return userfield7; }
            set { userfield7 = value; }
        }

        [Category("User-defined Attributes")]
        public string Userfield8
        {
            get { return userfield8; }
            set { userfield8 = value; }
        }

        [Category("User-defined Attributes")]
        public string DrawnBy
        {
            get { return drawnBy; }
            set { drawnBy = value; }
        }

        [Category("User-defined Attributes")]
        public string CheckedBy
        {
            get { return checkedBy; }
            set { checkedBy = value; }
        }

        #endregion

        DrawingHandler drawingHandler = new DrawingHandler();

        public void GetProperties()
        {
            Drawing drawing = drawingHandler.GetActiveDrawing();
            name = drawing.Name;
            title1 = drawing.Title1;
            title2 = drawing.Title2;
            title3 = drawing.Title3;
            height = drawing.Layout.SheetSize.Height.ToString();
            width = drawing.Layout.SheetSize.Width.ToString();
            drawing.GetUserProperty("DR_DRAWN_BY", ref drawnBy);
            drawing.GetUserProperty("DR_CHECKED_BY", ref checkedBy);
            drawing.GetUserProperty("DRAWING_USERFIELD_1", ref userfield1);
            drawing.GetUserProperty("DRAWING_USERFIELD_2", ref userfield2);
            drawing.GetUserProperty("DRAWING_USERFIELD_3", ref userfield3);
            drawing.GetUserProperty("DRAWING_USERFIELD_4", ref userfield4);
            drawing.GetUserProperty("DRAWING_USERFIELD_5", ref userfield5);
            drawing.GetUserProperty("DRAWING_USERFIELD_6", ref userfield6);
            drawing.GetUserProperty("DRAWING_USERFIELD_7", ref userfield7);
            drawing.GetUserProperty("DRAWING_USERFIELD_8", ref userfield8);
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            if (label == "Name") drawing.Name = name;
            if (label == "Title1") drawing.Title1 = title1;
            if (label == "Title2") drawing.Title2 = title2;
            if (label == "Title3") drawing.Title3 = title3;
            if (label == "DrawnBy") drawing.SetUserProperty("DR_DRAWN_BY", drawnBy);
            if (label == "CheckedBy") drawing.SetUserProperty("DR_CHECKED_BY", checkedBy);
            if (label == "Userfield1") drawing.SetUserProperty("DRAWING_USERFIELD_1", userfield1);
            if (label == "Userfield2") drawing.SetUserProperty("DRAWING_USERFIELD_2", userfield2);
            if (label == "Userfield3") drawing.SetUserProperty("DRAWING_USERFIELD_3", userfield3);
            if (label == "Userfield4") drawing.SetUserProperty("DRAWING_USERFIELD_4", userfield4);
            if (label == "Userfield5") drawing.SetUserProperty("DRAWING_USERFIELD_5", userfield5);
            if (label == "Userfield6") drawing.SetUserProperty("DRAWING_USERFIELD_6", userfield6);
            if (label == "Userfield7") drawing.SetUserProperty("DRAWING_USERFIELD_7", userfield7);
            if (label == "Userfield8") drawing.SetUserProperty("DRAWING_USERFIELD_8", userfield8);

            drawing.Modify();
            drawing.CommitChanges();
        }
    }
}
