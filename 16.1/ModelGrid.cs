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
    class ModelGrid
    {
        Model model = new Model();

        private string coordinateX;
        private string coordinateY;
        private string coordinateZ;
        private string labelX;
        private string labelY;
        private string labelZ;

        [Category("Coordinates")]
        public string CoordinateX
        {
            get { return coordinateX; }
            set { coordinateX = value; }
        }

        [Category("Coordinates")]
        public string CoordinateY
        {
            get { return coordinateY; }
            set { coordinateY = value; }
        }

        [Category("Coordinates")]
        public string CoordinateZ
        {
            get { return coordinateZ; }
            set { coordinateZ = value; }
        }

        [Category("Labels")]
        public string LabelX
        {
            get { return labelX; }
            set { labelX = value; }
        }

        [Category("Labels")]
        public string LabelY
        {
            get { return labelY; }
            set { labelY = value; }
        }

        [Category("Labels")]
        public string LabelZ
        {
            get { return labelZ; }
            set { labelZ = value; }
        }

        public void GetProperties()
        {
            ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
            if (modelObjectEnum.GetSize() == 1)
            {
                while (modelObjectEnum.MoveNext())
                {
                    if (modelObjectEnum.Current is Tekla.Structures.Model.Grid)
                    {
                        Grid grid = (Grid)modelObjectEnum.Current;

                        coordinateX = grid.CoordinateX;
                        coordinateY = grid.CoordinateY;
                        coordinateZ = grid.CoordinateZ;
                        labelX = grid.LabelX;
                        labelY = grid.LabelY;
                        labelZ = grid.LabelZ;
                    }
                }
            }
            if (modelObjectEnum.GetSize() > 1)
            {
                coordinateX = "";
                coordinateY = "";
                coordinateZ = "";
                labelX = "";
                labelY = "";
                labelZ = "";
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;

            ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
            while (modelObjectEnum.MoveNext())
            {
                if (modelObjectEnum.Current is Grid)
                {
                    Grid Grid = (Grid)modelObjectEnum.Current;

                    if (label == "CoordinateX") Grid.CoordinateX = coordinateX;
                    if (label == "CoordinateY") Grid.CoordinateY = coordinateY;
                    if (label == "CoordinateZ") Grid.CoordinateZ = coordinateZ;
                    if (label == "LabelX") Grid.LabelX = labelX;
                    if (label == "LabelY") Grid.LabelY = labelY;
                    if (label == "LabelZ") Grid.LabelZ = labelZ;

                    Grid.Modify();
                }
            }
        }
    }
}
