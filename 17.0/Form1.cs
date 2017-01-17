using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Tekla.Structures.Drawing;
using Geometry3d = Tekla.Structures.Geometry3d;

namespace TeklaProperties
{
    public partial class Form1 : Form
    {
        TeklaProperties.Properties.Settings settings = new TeklaProperties.Properties.Settings();
        private Tekla.Structures.Model.Events tsEvents = new Tekla.Structures.Model.Events();
        private Tekla.Structures.Drawing.UI.Events tsDrgEvents = new Tekla.Structures.Drawing.UI.Events();
        private delegate void UpdatePropertiesCallback();
        
        Model model = new Model();
        DrawingHandler drawingHandler = new DrawingHandler();

        ActiveDrawing activeDrawing = new ActiveDrawing();
        DrawingGraphic.Open.Line line = new DrawingGraphic.Open.Line();
        DrawingGraphic.Open.Arc arc = new DrawingGraphic.Open.Arc();
        DrawingGraphic.Open.Polyline polyline = new DrawingGraphic.Open.Polyline();
        DrawingGraphic.Closed.Rectangle rectangle = new DrawingGraphic.Closed.Rectangle();
        DrawingGraphic.Closed.Circle circle = new DrawingGraphic.Closed.Circle();
        DrawingGraphic.Closed.Polygon polygon = new DrawingGraphic.Closed.Polygon();
        DrawingText text = new DrawingText();
        DrawingView view = new DrawingView();
        DrawingSymbol symbol = new DrawingSymbol();
        DrawingMark mark = new DrawingMark();
        DrawingGrid drawingGrid = new DrawingGrid();
        DrawingPart drawingPart = new DrawingPart();
        DrawingBolt drawingBolt = new DrawingBolt();
        DrawingWeld drawingWeld = new DrawingWeld();
        DrawingDimension.Straight straightDim = new DrawingDimension.Straight();
        DrawingDimension.Radius radiusDim = new DrawingDimension.Radius();
        DrawingDimension.Radial radialDim = new DrawingDimension.Radial();
        DrawingDimension.Angle angleDim = new DrawingDimension.Angle();

        ModelPart.Beam beam = new ModelPart.Beam();
        ModelPart.ContourPlate contourPlate = new ModelPart.ContourPlate();
        ModelBolt modelBolt = new ModelBolt();
        ModelWeld modelWeld = new ModelWeld();
        ModelGrid modelGrid = new ModelGrid();
        ModelConnection modelConnection = new ModelConnection();
        ModelView modelView = new ModelView();

        public Form1()
        {
            
            InitializeComponent();
            this.Location = settings.formLocation;
            this.Size = settings.formSize;

            try
            {
                tsEvents.SelectionChange += new Tekla.Structures.Model.Events.SelectionChangeDelegate(tsEvents_SelectionChange);
                tsEvents.Register();

                tsDrgEvents.DrawingEditorOpened += new Tekla.Structures.Drawing.UI.Events.DrawingEditorOpenedDelegate(tsDrgEvents_DrawingEditorOpened);
                tsDrgEvents.DrawingEditorClosed += new Tekla.Structures.Drawing.UI.Events.DrawingEditorClosedDelegate(tsDrgEvents_DrawingEditorClosed);
                tsDrgEvents.SelectionChange += new Tekla.Structures.Drawing.UI.Events.SelectionChangeDelegate(tsDrgEvents_SelectionChange);
                tsDrgEvents.Register();
            }
            catch (InvalidOperationException e)
            {
                if (e.Source == "Tekla.Structures.Model" || e.Source == "Tekla.Structures.Drawing")
                {
                    UpdateProperties();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateProperties();
        }

        void tsDrgEvents_DrawingEditorClosed()
        {
            lock (this) UpdateProperties();
        }

        void tsDrgEvents_DrawingEditorOpened()
        {
            lock (this) UpdateProperties();
        }

        private void tsEvents_SelectionChange()
        {
            lock (this) UpdateProperties();
        }

        private void tsDrgEvents_SelectionChange()
        {
            lock (this) UpdateProperties();
        }

        private void UpdateProperties()
        {
            if (this.propertyGrid1.InvokeRequired)
            {
                UpdatePropertiesCallback d = new UpdatePropertiesCallback(UpdateProperties);
                this.Invoke(d);
            }
            else
            {
                Drawing drawing = drawingHandler.GetActiveDrawing();
                if (drawing != null)
                {
                    DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    statusLabel.Text = drawingObjectEnum.GetSize().ToString() + " objects selected";
                    ArrayList arraySelectedObjects = new ArrayList();
                    if (drawingObjectEnum.GetSize() == 0)
                    {
                        activeDrawing.GetProperties();
                        propertyGrid1.SelectedObject = activeDrawing;
                    }

                    while (drawingObjectEnum.MoveNext())
                    {
                        string objectType = drawingObjectEnum.Current.GetType().ToString();
                        if (!arraySelectedObjects.Contains(objectType)) arraySelectedObjects.Add(objectType);
                    }

                    comboBox1.Text = "";
                    comboBox1.Items.Clear();

                    arraySelectedObjects.Sort();
                    foreach (string objectName in arraySelectedObjects)
                        comboBox1.Items.Add(objectName);

                    if (arraySelectedObjects.Count == 1)
                    {

                        if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Part")
                        {
                            drawingPart.GetProperties();
                            propertyGrid1.SelectedObject = drawingPart;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Weld")
                        {
                            drawingWeld.GetProperties();
                            propertyGrid1.SelectedObject = drawingWeld;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.View")
                        {
                            view.GetProperties();
                            propertyGrid1.SelectedObject = view;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Bolt")
                        {
                            drawingBolt.GetProperties();
                            propertyGrid1.SelectedObject = drawingBolt;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.StraightDimensionSet")
                        {
                            straightDim.GetProperties();
                            propertyGrid1.SelectedObject = straightDim;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.RadiusDimension")
                        {
                            radiusDim.GetProperties();
                            propertyGrid1.SelectedObject = radiusDim;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.CurvedDimensionSetRadial")
                        {
                            radialDim.GetProperties();
                            propertyGrid1.SelectedObject = radialDim;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.AngleDimension")
                        {
                            angleDim.GetProperties();
                            propertyGrid1.SelectedObject = angleDim;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.GridLine")
                        {
                            drawingGrid.GetProperties();
                            propertyGrid1.SelectedObject = drawingGrid;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Line")
                        {
                            line.GetProperties();
                            propertyGrid1.SelectedObject = line;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Arc")
                        {
                            arc.GetProperties();
                            propertyGrid1.SelectedObject = arc;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Polyline")
                        {
                            polyline.GetProperties();
                            propertyGrid1.SelectedObject = polyline;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Rectangle")
                        {
                            rectangle.GetProperties();
                            propertyGrid1.SelectedObject = rectangle;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Circle")
                        {
                            circle.GetProperties();
                            propertyGrid1.SelectedObject = circle;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Polygon")
                        {
                            polygon.GetProperties();
                            propertyGrid1.SelectedObject = polygon;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Text")
                        {
                            text.GetProperties();
                            propertyGrid1.SelectedObject = text;
                        }
                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Symbol")
                        {
                            symbol.GetProperties();
                            propertyGrid1.SelectedObject = symbol;
                        }
                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Drawing.Mark")
                        {
                            mark.GetProperties();
                            propertyGrid1.SelectedObject = mark;
                        }

                        else
                        {
                            activeDrawing.GetProperties();
                            propertyGrid1.SelectedObject = activeDrawing;
                        }

                        comboBox1.SelectedText = arraySelectedObjects[0].ToString();
                        propertyGrid1.Focus();
                    }

                    if (arraySelectedObjects.Count > 1) propertyGrid1.SelectedObject = null;

                }

                if (drawing == null)
                {
                    Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    ModelObjectEnumerator modelObjectEnum = modelObjectSelector.GetSelectedObjects();
                    statusLabel.Text = modelObjectEnum.GetSize().ToString() + " objects selected";
                    ArrayList arraySelectedObjects = new ArrayList();

                    if (modelObjectEnum.GetSize() == 0)
                        propertyGrid1.SelectedObject = null;

                    while (modelObjectEnum.MoveNext())
                    {
                        string objectType = modelObjectEnum.Current.GetType().ToString();
                        if (!arraySelectedObjects.Contains(objectType)) arraySelectedObjects.Add(objectType);
                    }

                    comboBox1.Text = "";
                    comboBox1.Items.Clear();

                    arraySelectedObjects.Sort();
                    foreach (string objectName in arraySelectedObjects)
                        comboBox1.Items.Add(objectName);

                    if (arraySelectedObjects.Count == 0)
                    {
                        Tekla.Structures.Model.UI.ModelViewEnumerator modelViewEnum = Tekla.Structures.Model.UI.ViewHandler.GetSelectedViews();
                        if (modelViewEnum.Count == 1)
                        {
                            modelView.GetProperties();
                            propertyGrid1.SelectedObject = modelView;

                            comboBox1.SelectedText = "Tekla.Structures.Model.UI.View";
                            propertyGrid1.Focus();
                        }
                    }

                    if (arraySelectedObjects.Count == 1)
                    {
                        if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Model.Beam")
                        {
                            beam.GetProperties();
                            propertyGrid1.SelectedObject = beam;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Model.ContourPlate")
                        {
                            contourPlate.GetProperties();
                            propertyGrid1.SelectedObject = contourPlate;
                        }

                        else if (arraySelectedObjects[0].ToString().Contains("Tekla.Structures.Model.Weld"))
                        {
                            modelWeld.GetProperties();
                            propertyGrid1.SelectedObject = modelWeld;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Model.PolygonWeld")
                        {
                            modelWeld.GetProperties();
                            propertyGrid1.SelectedObject = modelWeld;
                        }

                        else if (arraySelectedObjects[0].ToString().Contains("Tekla.Structures.Model.Bolt"))
                        {
                            modelBolt.GetProperties();
                            propertyGrid1.SelectedObject = modelBolt;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Model.Grid")
                        {
                            modelGrid.GetProperties();
                            propertyGrid1.SelectedObject = modelGrid;
                        }

                        else if (arraySelectedObjects[0].ToString() == "Tekla.Structures.Model.Connection")
                        {
                            modelConnection.GetProperties();
                            propertyGrid1.SelectedObject = modelConnection;
                        }

                        comboBox1.SelectedText = arraySelectedObjects[0].ToString();
                        propertyGrid1.Focus();
                    }

                    if (arraySelectedObjects.Count > 1) propertyGrid1.SelectedObject = null;
                }
            }
        }

        ArrayList GetModelObjectsByType(string ObjectType)
        {
            ArrayList ObjectsToBeSelected = new ArrayList();

            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            foreach (Tekla.Structures.Model.ModelObject modelObject in modelObjectSelector.GetSelectedObjects())
            {
                if (modelObject.GetType().ToString() == ObjectType)
                    ObjectsToBeSelected.Add(modelObject);
            }
            return ObjectsToBeSelected;
        }

        ArrayList GetDrawingObjectsByType(string ObjectType)
        {
            ArrayList ObjectsToBeSelected = new ArrayList();
            foreach (DrawingObject drawingObject in drawingHandler.GetDrawingObjectSelector().GetSelected())
            {
                if (drawingObject.GetType().ToString() == ObjectType)
                    ObjectsToBeSelected.Add(drawingObject);
            }
            return ObjectsToBeSelected;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            Drawing drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null)
            {
                Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                ArrayList arrayObjectsSelected = GetModelObjectsByType(comboBox1.Text);
                if (arrayObjectsSelected.Count > 0)
                    modelObjectSelector.Select(arrayObjectsSelected);
            }
            if (drawing != null)
            {
                ArrayList arrayObjectsSelected = GetDrawingObjectsByType(comboBox1.Text);
                if (arrayObjectsSelected.Count > 0)
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayObjectsSelected, false);
            }
        }

        private void propertyGrid1_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            Drawing drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null)
            {
                if (comboBox1.Text == "Tekla.Structures.Model.Beam") beam.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Model.ContourPlate") contourPlate.Modify(e);
                if (comboBox1.Text.Contains("Tekla.Structures.Model.Bolt")) modelBolt.Modify(e);
                if (comboBox1.Text.Contains("Tekla.Structures.Model.Weld")) modelWeld.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Model.Grid") modelGrid.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Model.UI.View") modelView.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Model.Connection") modelConnection.Modify(e);

                model.CommitChanges();
            }
            if (drawing != null)
            {
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();

                if (drawingObjectEnum.GetSize() == 0) activeDrawing.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Line") line.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Arc") arc.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Polyline") polyline.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Rectangle") rectangle.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Circle") circle.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Polygon") polygon.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Text") text.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.View") view.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.GridLine") drawingGrid.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Part") drawingPart.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Bolt") drawingBolt.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Symbol") symbol.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.Mark") mark.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.StraightDimensionSet") straightDim.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.RadiusDimension") radiusDim.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.CurvedDimensionSetRadial") radialDim.Modify(e);
                if (comboBox1.Text == "Tekla.Structures.Drawing.AngleDimension") angleDim.Modify(e);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            tsEvents.UnRegister();
            tsDrgEvents.UnRegister();

            settings.formLocation = this.Location;
            settings.formSize = this.Size;
            settings.Save();
        }
    }
}