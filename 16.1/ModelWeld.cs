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
    class ModelWeld
    {
        Model model = new Model();

        public class WeldTypeGridEditor : UITypeEditor
        {
            public override bool GetPaintValueSupported(ITypeDescriptorContext context)
            {
                return true;
            }

            public override void PaintValue(PaintValueEventArgs e)
            {
                string m = this.GetType().Module.Name;
                m = m.Substring(0, m.Length - 4);
                ResourceManager resourceManager = new ResourceManager("TeklaProperties.Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());

                int i = (int)e.Value;
                string _SourceName = "";
                switch (i)
                {
                    case ((int)WeldType.WELD_TYPE_NONE):
                        _SourceName = "WELD_TYPE_NONE";
                        break;
                    case ((int)WeldType.WELD_TYPE_FILLET):
                        _SourceName = "WELD_TYPE_FILLET";
                        break;
                    case ((int)WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT):
                        _SourceName = "WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT):
                        _SourceName = "WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT):
                        _SourceName = "WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE):
                        _SourceName = "WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE";
                        break;
                    case ((int)WeldType.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE):
                        _SourceName = "WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE";
                        break;
                    case ((int)WeldType.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT):
                        _SourceName = "WELD_TYPE_U_GROOVE_SINGLE_U_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_J_GROOVE_J_BUTT):
                        _SourceName = "WELD_TYPE_J_GROOVE_J_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_FLARE_V_GROOVE):
                        _SourceName = "WELD_TYPE_FLARE_V_GROOVE";
                        break;
                    case ((int)WeldType.WELD_TYPE_FLARE_BEVEL_GROOVE):
                        _SourceName = "WELD_TYPE_FLARE_BEVEL_GROOVE";
                        break;
                    case ((int)WeldType.WELD_TYPE_EDGE_FLANGE):
                        _SourceName = "WELD_TYPE_EDGE_FLANGE";
                        break;
                    case ((int)WeldType.WELD_TYPE_CORNER_FLANGE):
                        _SourceName = "WELD_TYPE_CORNER_FLANGE";
                        break;
                    case ((int)WeldType.WELD_TYPE_PLUG):
                        _SourceName = "WELD_TYPE_PLUG";
                        break;
                    case ((int)WeldType.WELD_TYPE_BEVEL_BACKING):
                        _SourceName = "WELD_TYPE_BEVEL_BACKING";
                        break;
                    case ((int)WeldType.WELD_TYPE_SPOT):
                        _SourceName = "WELD_TYPE_SPOT";
                        break;
                    case ((int)WeldType.WELD_TYPE_SEAM):
                        _SourceName = "WELD_TYPE_SEAM";
                        break;
                    case ((int)WeldType.WELD_TYPE_SLOT):
                        _SourceName = "WELD_TYPE_SLOT";
                        break;
                    case ((int)WeldType.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET):
                        _SourceName = "WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET";
                        break;
                    case ((int)WeldType.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET):
                        _SourceName = "WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET";
                        break;
                    case ((int)WeldType.WELD_TYPE_MELT_THROUGH):
                        _SourceName = "WELD_TYPE_MELT_THROUGH";
                        break;
                    case ((int)WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT):
                        _SourceName = "STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT";
                        break;
                    case ((int)WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT):
                        _SourceName = "STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT";
                        break;
                    case ((int)WeldType.WELD_TYPE_EDGE):
                        _SourceName = "WELD_TYPE_EDGE";
                        break;
                    case ((int)WeldType.WELD_TYPE_ISO_SURFACING):
                        _SourceName = "WELD_TYPE_ISO_SURFACING";
                        break;
                    case ((int)WeldType.WELD_TYPE_FOLD):
                        _SourceName = "WELD_TYPE_FOLD";
                        break;
                    case ((int)WeldType.WELD_TYPE_INCLINED):
                        _SourceName = "WELD_TYPE_INCLINED";
                        break;
                }

                //Draw the corresponding image
                Bitmap newImage = (Bitmap)resourceManager.GetObject(_SourceName);
                Rectangle destRect = e.Bounds;
                destRect.Width = newImage.Width;
                newImage.MakeTransparent();
                e.Graphics.DrawImage(newImage, destRect);
            }
        }

        public class WeldContourTypeGridEditor : UITypeEditor
        {

        }

        private string sizeAbove;
        private WeldType weldTypeAbove;
        private ContourType contourTypeAbove;
        private string sizeBelow;
        private WeldType weldTypeBelow;
        private ContourType contourTypeBelow;
        private string refText;
        private Bool around;

        [Category("Above Line")]
        public string SizeAbove
        {
            get { return sizeAbove; }
            set { sizeAbove = value; }
        }

        [Category("Above Line")]
        [Editor(typeof(WeldTypeGridEditor), typeof(UITypeEditor))]
        public WeldType WeldTypeAbove
        {
            get { return weldTypeAbove; }
            set { weldTypeAbove = value; }
        }

        [Category("Above Line")]
        [Editor(typeof(WeldContourTypeGridEditor), typeof(UITypeEditor))]
        public ContourType ContourTypeAbove
        {
            get { return contourTypeAbove; }
            set { contourTypeAbove = value; }
        }

        [Category("Below Line")]
        public string SizeBelow
        {
            get { return sizeBelow; }
            set { sizeBelow = value; }
        }

        [Category("Below Line")]
        [Editor(typeof(WeldTypeGridEditor), typeof(UITypeEditor))]
        public WeldType WeldTypeBelow
        {
            get { return weldTypeBelow; }
            set { weldTypeBelow = value; }
        }

        [Category("Below Line")]
        [Editor(typeof(WeldContourTypeGridEditor), typeof(UITypeEditor))]
        public ContourType ContourTypeBelow
        {
            get { return contourTypeBelow; }
            set { contourTypeBelow = value; }
        }

        [Category("Common Attributes")]
        public string RefText
        {
            get { return refText; }
            set { refText = value; }
        }

        [Category("Common Attributes")]
        public Bool Around
        {
            get { return around; }
            set { around = value; }
        }

        public void GetProperties()
        {
            Model model = new Model();
            ModelObjectEnumerator modelobjenum = model.GetModelObjectSelector().GetSelectedObjects();
            if (modelobjenum.GetSize() == 1)
            {
                while (modelobjenum.MoveNext())
                {
                    if (modelobjenum.Current is Tekla.Structures.Model.BaseWeld)
                    {
                        BaseWeld weld = (BaseWeld)modelobjenum.Current;

                        refText = weld.ReferenceText;
                        if (weld.AroundWeld) around = Bool.True; else around = Bool.False;

                        sizeAbove = weld.SizeAbove.ToString();

                        if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_NONE) weldTypeAbove = WeldType.WELD_TYPE_NONE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET) weldTypeAbove = WeldType.WELD_TYPE_FILLET;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT) weldTypeAbove = WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weldTypeAbove = WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT) weldTypeAbove = WeldType.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE) weldTypeAbove = WeldType.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE) weldTypeAbove = WeldType.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT) weldTypeAbove = WeldType.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_J_GROOVE_J_BUTT) weldTypeAbove = WeldType.WELD_TYPE_J_GROOVE_J_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_V_GROOVE) weldTypeAbove = WeldType.WELD_TYPE_FLARE_V_GROOVE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_BEVEL_GROOVE) weldTypeAbove = WeldType.WELD_TYPE_FLARE_BEVEL_GROOVE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE_FLANGE) weldTypeAbove = WeldType.WELD_TYPE_EDGE_FLANGE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_CORNER_FLANGE) weldTypeAbove = WeldType.WELD_TYPE_CORNER_FLANGE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_PLUG) weldTypeAbove = WeldType.WELD_TYPE_PLUG;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_BACKING) weldTypeAbove = WeldType.WELD_TYPE_BEVEL_BACKING;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SPOT) weldTypeAbove = WeldType.WELD_TYPE_SPOT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SEAM) weldTypeAbove = WeldType.WELD_TYPE_SEAM;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_SLOT) weldTypeAbove = WeldType.WELD_TYPE_SLOT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET) weldTypeAbove = WeldType.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET) weldTypeAbove = WeldType.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_MELT_THROUGH) weldTypeAbove = WeldType.WELD_TYPE_MELT_THROUGH;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT) weldTypeAbove = WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weldTypeAbove = WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE) weldTypeAbove = WeldType.WELD_TYPE_EDGE;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_ISO_SURFACING) weldTypeAbove = WeldType.WELD_TYPE_ISO_SURFACING;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_FOLD) weldTypeAbove = WeldType.WELD_TYPE_FOLD;
                        else if (weld.TypeAbove == BaseWeld.WeldTypeEnum.WELD_TYPE_INCLINED) weldTypeAbove = WeldType.WELD_TYPE_INCLINED;
                        else Console.WriteLine(weld.TypeAbove);

                        if (weld.ContourAbove == BaseWeld.WeldContourEnum.WELD_CONTOUR_NONE) contourTypeAbove = ContourType.None;
                        else if (weld.ContourAbove == BaseWeld.WeldContourEnum.WELD_CONTOUR_FLUSH) contourTypeAbove = ContourType.Flush;

                        sizeBelow = weld.SizeBelow.ToString();

                        if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_NONE) weldTypeBelow = WeldType.WELD_TYPE_NONE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET) weldTypeBelow = WeldType.WELD_TYPE_FILLET;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT) weldTypeBelow = WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weldTypeBelow = WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT) weldTypeBelow = WeldType.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE) weldTypeBelow = WeldType.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE) weldTypeBelow = WeldType.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT) weldTypeBelow = WeldType.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_J_GROOVE_J_BUTT) weldTypeBelow = WeldType.WELD_TYPE_J_GROOVE_J_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_V_GROOVE) weldTypeBelow = WeldType.WELD_TYPE_FLARE_V_GROOVE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_BEVEL_GROOVE) weldTypeBelow = WeldType.WELD_TYPE_FLARE_BEVEL_GROOVE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE_FLANGE) weldTypeBelow = WeldType.WELD_TYPE_EDGE_FLANGE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_CORNER_FLANGE) weldTypeBelow = WeldType.WELD_TYPE_CORNER_FLANGE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_PLUG) weldTypeBelow = WeldType.WELD_TYPE_PLUG;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_BACKING) weldTypeBelow = WeldType.WELD_TYPE_BEVEL_BACKING;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SPOT) weldTypeBelow = WeldType.WELD_TYPE_SPOT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SEAM) weldTypeBelow = WeldType.WELD_TYPE_SEAM;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_SLOT) weldTypeBelow = WeldType.WELD_TYPE_SLOT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET) weldTypeBelow = WeldType.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET) weldTypeBelow = WeldType.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_MELT_THROUGH) weldTypeBelow = WeldType.WELD_TYPE_MELT_THROUGH;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT) weldTypeBelow = WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weldTypeBelow = WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE) weldTypeBelow = WeldType.WELD_TYPE_EDGE;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_ISO_SURFACING) weldTypeBelow = WeldType.WELD_TYPE_ISO_SURFACING;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_FOLD) weldTypeBelow = WeldType.WELD_TYPE_FOLD;
                        else if (weld.TypeBelow == BaseWeld.WeldTypeEnum.WELD_TYPE_INCLINED) weldTypeBelow = WeldType.WELD_TYPE_INCLINED;
                        else Console.WriteLine(weld.TypeBelow);

                        if (weld.ContourBelow == BaseWeld.WeldContourEnum.WELD_CONTOUR_NONE) contourTypeBelow = ContourType.None;
                        else if (weld.ContourBelow == BaseWeld.WeldContourEnum.WELD_CONTOUR_FLUSH) contourTypeBelow = ContourType.Flush;
                    }
                }
            }
            else
            {
                sizeAbove = "";
                weldTypeAbove = new WeldType();
                contourTypeAbove = new ContourType();
                sizeBelow = "";
                weldTypeBelow = new WeldType();
                contourTypeBelow = new ContourType();
                refText = "";
                around = new Bool();
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
            while (modelObjectEnum.MoveNext())
            {
                if (modelObjectEnum.Current is BaseWeld)
                {
                    BaseWeld weld = (BaseWeld)modelObjectEnum.Current;
                    if (label == "SizeAbove") weld.SizeAbove = double.Parse(sizeAbove);

                    if (label == "WeldTypeAbove")
                    {
                        if (weldTypeAbove == WeldType.WELD_TYPE_NONE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_NONE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_FILLET) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_J_GROOVE_J_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_J_GROOVE_J_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_FLARE_V_GROOVE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_V_GROOVE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_FLARE_BEVEL_GROOVE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_BEVEL_GROOVE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_EDGE_FLANGE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE_FLANGE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_CORNER_FLANGE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_CORNER_FLANGE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_PLUG) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_PLUG;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_BEVEL_BACKING) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_BACKING;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SPOT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SPOT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SEAM) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SEAM;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_SLOT) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_SLOT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_MELT_THROUGH) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_MELT_THROUGH;
                        else if (weldTypeAbove == WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weldTypeAbove == WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weld.TypeAbove = BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_EDGE) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_ISO_SURFACING) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_ISO_SURFACING;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_FOLD) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_FOLD;
                        else if (weldTypeAbove == WeldType.WELD_TYPE_INCLINED) weld.TypeAbove = BaseWeld.WeldTypeEnum.WELD_TYPE_INCLINED;
                    }
                    if (label == "ContourTypeAbove")
                    {
                        if (contourTypeAbove == ContourType.None) weld.ContourAbove = BaseWeld.WeldContourEnum.WELD_CONTOUR_NONE;
                        else if (contourTypeAbove == ContourType.Flush) weld.ContourAbove = BaseWeld.WeldContourEnum.WELD_CONTOUR_FLUSH;
                    }

                    if (label == "SizeBelow") weld.SizeBelow = double.Parse(sizeBelow);

                    if (label == "WeldTypeBelow")
                    {
                        if (weldTypeBelow == WeldType.WELD_TYPE_NONE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_NONE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_FILLET) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_FILLET;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_U_GROOVE_SINGLE_U_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_J_GROOVE_J_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_J_GROOVE_J_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_FLARE_V_GROOVE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_V_GROOVE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_FLARE_BEVEL_GROOVE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_FLARE_BEVEL_GROOVE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_EDGE_FLANGE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE_FLANGE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_CORNER_FLANGE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_CORNER_FLANGE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_PLUG) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_PLUG;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_BEVEL_BACKING) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_BEVEL_BACKING;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SPOT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SPOT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SEAM) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SEAM;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_SLOT) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_SLOT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_MELT_THROUGH) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_MELT_THROUGH;
                        else if (weldTypeBelow == WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT;
                        else if (weldTypeBelow == WeldType.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT) weld.TypeBelow = BaseWeld.WeldTypeEnum.STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_EDGE) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_EDGE;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_ISO_SURFACING) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_ISO_SURFACING;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_FOLD) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_FOLD;
                        else if (weldTypeBelow == WeldType.WELD_TYPE_INCLINED) weld.TypeBelow = BaseWeld.WeldTypeEnum.WELD_TYPE_INCLINED;
                    }
                    if (label == "ContourTypeBelow")
                    {
                        if (contourTypeBelow == ContourType.None) weld.ContourBelow = BaseWeld.WeldContourEnum.WELD_CONTOUR_NONE;
                        else if (contourTypeBelow == ContourType.Flush) weld.ContourBelow = BaseWeld.WeldContourEnum.WELD_CONTOUR_FLUSH;
                    }

                    if (label == "RefText") weld.ReferenceText = refText;

                    if (label == "Around")
                    {
                        if (around == Bool.True) weld.AroundWeld = true; else weld.AroundWeld = false;
                    }

                    weld.Modify();
                }
            }
        }
    }
}
