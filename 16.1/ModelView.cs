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
using Tekla.Structures.Geometry3d;
using System.Drawing.Design;
using System.Resources;
using Reflection = System.Reflection;

namespace TeklaProperties
{
    class ModelView
    {
        Model model = new Model();

        private string name;
        private string viewDepthUp;
        private string viewDepthDown;
        private string viewRendering;

        [Category("View")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [Category("View")]
        public string ViewDepthUp
        {
            get { return viewDepthUp; }
            set { viewDepthUp = value; }
        }

        [Category("View")]
        public string ViewDepthDown
        {
            get { return viewDepthDown; }
            set { viewDepthDown = value; }
        }

        public void GetProperties()
        {
            Tekla.Structures.Model.UI.ModelViewEnumerator modelViewEnum = Tekla.Structures.Model.UI.ViewHandler.GetSelectedViews();

            if (modelViewEnum.Count == 1)
            {
                while (modelViewEnum.MoveNext())
                {
                    Tekla.Structures.Model.UI.View view = (Tekla.Structures.Model.UI.View)modelViewEnum.Current;
                    name = view.Name;
                    viewDepthUp = view.ViewDepthUp.ToString("F0");
                    viewDepthDown = view.ViewDepthDown.ToString("F0");
                }
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Tekla.Structures.Model.UI.ModelViewEnumerator modelViewEnum = Tekla.Structures.Model.UI.ViewHandler.GetSelectedViews();
            if (modelViewEnum.Count == 1)
            {
                while (modelViewEnum.MoveNext())
                {
                    Tekla.Structures.Model.UI.View view = (Tekla.Structures.Model.UI.View)modelViewEnum.Current;
                    if (label == "Name") view.Name = name;
                    if (label == "ViewDepthUp") view.ViewDepthUp = double.Parse(viewDepthUp);
                    if (label == "ViewDepthDown") view.ViewDepthDown = double.Parse(viewDepthDown);
                    
                    view.Modify();
                    model.CommitChanges();
                }
            }
        }
    }
}

