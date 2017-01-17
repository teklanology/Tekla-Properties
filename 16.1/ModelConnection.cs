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
    class ModelConnection
    {
        Model model = new Model();

        private string code;
        private string name;
        private string number;

        [Category("Connection")]
        public string Code
        {
            get { return code; }
            set { code = value; }
        }

        [Category("Connection")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [Category("Connection")]
        public string Number
        {
            get { return number; }
            set { number = value; }
        }


        public void GetProperties()
        {
            ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
            if (modelObjectEnum.GetSize() == 1)
            {
                while (modelObjectEnum.MoveNext())
                {
                    if (modelObjectEnum.Current is Tekla.Structures.Model.Connection)
                    {
                        Connection connection = (Connection)modelObjectEnum.Current;
                        code = connection.Code;
                        name = connection.Name;
                        //connection.Number);
                        //connection.SetAttribute
                    }
                }
            }
            if (modelObjectEnum.GetSize() > 1)
            {
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;

            ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
            while (modelObjectEnum.MoveNext())
            {
                if (modelObjectEnum.Current is Connection)
                {
                    Connection connection = (Connection)modelObjectEnum.Current;
                }
            }
        }
    }
}
