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
    class ModelBolt
    {
        Model model = new Model();

        private string boltSize;
        private string boltStandard;
        private BoltTypeEnum boltType;
        private ThreadMaterialEnum threadMaterial;
        private string cutLength;
        private string extraLength;
        private string boltGroupShape;
        private string boltDistX;
        private string boltDistY;
        private string tolerance;
        private string slottedHoleX;
        private string slottedHoleY;
        private RotateSlotEnum rotateSlots;
        private HoleTypeEnum holeType;
        private PositionPlaneEnum positionOnPlane;
        private string positionOnPlaneOffset;
        private PositionRotationEnum positionRotation;
        private string positionRotationOffset;
        private PositionDepthEnum positionAtDepth;
        private string positionAtDepthOffset;
        private string offsetFromStartX;
        private string offsetFromStartY;
        private string offsetFromStartZ;
        private string offsetFromEndX;
        private string offsetFromEndY;
        private string offsetFromEndZ;
        private Bool hole1;
        private Bool hole2;
        private Bool hole3;
        private Bool hole4;
        private Bool hole5;
        private Bool boltShaft;
        private Bool washer1;
        private Bool washer2;
        private Bool washer3;
        private Bool nut1;
        private Bool nut2;

        [Category("Bolt")]
        public string BoltSize
        {
            get { return boltSize; }
            set { boltSize = value; }
        }

        [Category("Bolt")]
        public string BoltStandard
        {
            get { return boltStandard; }
            set { boltStandard = value; }
        }

        [Category("Bolt")]
        public BoltTypeEnum BoltType
        {
            get { return boltType; }
            set { boltType = value; }
        }

        [Category("Bolt")]
        public ThreadMaterialEnum ThreadMaterial
        {
            get { return threadMaterial; }
            set { threadMaterial = value; }
        }

        [Category("Bolt")]
        public string CutLength
        {
            get { return cutLength; }
            set { cutLength = value; }
        }

        [Category("Bolt")]
        public string ExtraLength
        {
            get { return extraLength; }
            set { extraLength = value; }
        }

        [ReadOnly(true)]
        [Category("Bolt Group")]
        public string BoltGroupShape
        {
            get { return boltGroupShape; }
            set { boltGroupShape = value; }
        }

        [Category("Bolt Group")]
        [DisplayName(@"BoltDistX \ No Bolts")]
        public string BoltDistX
        {
            get { return boltDistX; }
            set { boltDistX = value; }
        }

        [Category("Bolt Group")]
        [DisplayName(@"BoltDistY \ Diameter")]
        public string BoltDistY
        {
            get { return boltDistY; }
            set { boltDistY = value; }
        }

        [Category("Hole")]
        public string Tolerance
        {
            get { return tolerance; }
            set { tolerance = value; }
        }

        [Category("Hole")]
        public HoleTypeEnum HoleType
        {
            get
            {
                //if (hole1 == Bool.False && hole2 == Bool.False && hole3 == Bool.False && hole4 == Bool.False && hole5 == Bool.False)
                    //holeType = HoleTypeEnum.Standard;
                return holeType;
            }
            set { holeType = value; }
        }

        [Category("Hole")]
        public string SlottedHoleX
        {
            get { return slottedHoleX; }
            set { slottedHoleX = value; }
        }

        [Category("Hole")]
        public string SlottedHoleY
        {
            get { return slottedHoleY; }
            set { slottedHoleY = value; }
        }

        [Category("Hole")]
        public RotateSlotEnum RotateSlots
        {
            get { return rotateSlots; }
            set { rotateSlots = value; }
        }

        [ReadOnly(true)]
        [Category("Position")]
        public PositionPlaneEnum PositionOnPlane
        {
            get { return positionOnPlane; }
            set { positionOnPlane = value; }
        }

        [Category("Position")]
        public string PositionOnPlaneOffset
        {
            get { return positionOnPlaneOffset; }
            set { positionOnPlaneOffset = value; }
        }

        [Category("Position")]
        public PositionRotationEnum PositionRotation
        {
            get { return positionRotation; }
            set { positionRotation = value; }
        }

        [Category("Position")]
        public string PositionRotationOffset
        {
            get { return positionRotationOffset; }
            set { positionRotationOffset = value; }
        }

        [ReadOnly(true)]
        [Category("Position")]
        public PositionDepthEnum PositionAtDepth
        {
            get { return positionAtDepth; }
            set { positionAtDepth = value; }
        }

        [Category("Position")]
        public string PositionAtDepthOffset
        {
            get { return positionAtDepthOffset; }
            set { positionAtDepthOffset = value; }
        }

        [Category("Offset")]
        public string OffsetFromStartX
        {
            get { return offsetFromStartX; }
            set { offsetFromStartX = value; }
        }

        [Category("Offset")]
        public string OffsetFromStartY
        {
            get { return offsetFromStartY; }
            set { offsetFromStartY = value; }
        }

        [Category("Offset")]
        public string OffsetFromStartZ
        {
            get { return offsetFromStartZ; }
            set { offsetFromStartZ = value; }
        }

        [Category("Offset")]
        public string OffsetFromEndX
        {
            get { return offsetFromEndX; }
            set { offsetFromEndX = value; }
        }

        [Category("Offset")]
        public string OffsetFromEndY
        {
            get { return offsetFromEndY; }
            set { offsetFromEndY = value; }
        }

        [Category("Offset")]
        public string OffsetFromEndZ
        {
            get { return offsetFromEndZ; }
            set { offsetFromEndZ = value; }
        }

        [Category("Parts with Slotted Holes")]
        public Bool Hole1
        {
            get { return hole1; }
            set { hole1 = value; }
        }

        [Category("Parts with Slotted Holes")]
        public Bool Hole2
        {
            get { return hole2; }
            set { hole2 = value; }
        }

        [Category("Parts with Slotted Holes")]
        public Bool Hole3
        {
            get { return hole3; }
            set { hole3 = value; }
        }

        [Category("Parts with Slotted Holes")]
        public Bool Hole4
        {
            get { return hole4; }
            set { hole4 = value; }
        }

        [Category("Parts with Slotted Holes")]
        public Bool Hole5
        {
            get { return hole5; }
            set { hole5 = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool BoltShaft
        {
            get { return boltShaft; }
            set { boltShaft = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool Washer1
        {
            get { return washer1; }
            set { washer1 = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool Washer2
        {
            get { return washer2; }
            set { washer2 = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool Washer3
        {
            get { return washer3; }
            set { washer3 = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool Nut1
        {
            get { return nut1; }
            set { nut1 = value; }
        }

        [Category("Include in Bolt Assembly")]
        public Bool Nut2
        {
            get { return nut2; }
            set { nut2 = value; }
        }


        public void GetProperties()
        {
            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnum = modelObjectSelector.GetSelectedObjects();
            if (modelObjectEnum.GetSize() == 1)
            {
                while (modelObjectEnum.MoveNext())
                {
                    if (modelObjectEnum.Current is Tekla.Structures.Model.BoltGroup)
                    {
                        BoltGroup boltGroup = (BoltGroup)modelObjectEnum.Current;
                        boltSize = boltGroup.BoltSize.ToString();
                        boltStandard = boltGroup.BoltStandard.ToString();
                        if (boltGroup.BoltType == BoltGroup.BoltTypeEnum.BOLT_TYPE_SITE)
                            boltType = BoltTypeEnum.Site;
                        else if (boltGroup.BoltType == BoltGroup.BoltTypeEnum.BOLT_TYPE_WORKSHOP)
                            boltType = BoltTypeEnum.Workshop;

                        if (boltGroup.ThreadInMaterial == BoltGroup.BoltThreadInMaterialEnum.THREAD_IN_MATERIAL_YES)
                            threadMaterial = ThreadMaterialEnum.Yes;
                        else if (boltGroup.ThreadInMaterial == BoltGroup.BoltThreadInMaterialEnum.THREAD_IN_MATERIAL_NO)
                            threadMaterial = ThreadMaterialEnum.No;

                        cutLength = boltGroup.CutLength.ToString();
                        extraLength = boltGroup.ExtraLength.ToString();
                        boltGroupShape = boltGroup.GetType().Name;

                        if (boltGroup is BoltArray)
                        {
                            BoltArray boltArray = (BoltArray)boltGroup;

                            string boltPositionX = "";
                            for (int i = 0; i < boltArray.GetBoltDistXCount(); i++)
                                boltPositionX = boltPositionX + boltArray.GetBoltDistX(i).ToString() + " ";
                            boltDistX = boltPositionX.Trim();

                            string boltPositionY = "";
                            for (int i = 0; i < boltArray.GetBoltDistYCount(); i++)
                                boltPositionY = boltPositionY + boltArray.GetBoltDistY(i).ToString() + " ";
                            boltDistY = boltPositionY.Trim();
                        }
                        else if (boltGroup is BoltXYList)
                        {
                            BoltXYList boltXYList = (BoltXYList)boltGroup;

                            string boltPositionX = "";
                            for (int i = 0; i < boltXYList.GetBoltDistXCount(); i++)
                                boltPositionX = boltPositionX + boltXYList.GetBoltDistX(i).ToString() + " ";
                            boltDistX = boltPositionX.Trim();

                            string boltPositionY = "";
                            for (int i = 0; i < boltXYList.GetBoltDistYCount(); i++)
                                boltPositionY = boltPositionY + boltXYList.GetBoltDistY(i).ToString() + " ";
                            boltDistY = boltPositionY.Trim();
                        }
                        else if (boltGroup is BoltCircle)
                        {
                            BoltCircle boltCircle = (BoltCircle)boltGroup;
                            boltDistX = boltCircle.NumberOfBolts.ToString();
                            boltDistY = boltCircle.Diameter.ToString();
                        }

                        tolerance = boltGroup.Tolerance.ToString();

                        if (boltGroup.HoleType == BoltGroup.BoltHoleTypeEnum.HOLE_TYPE_SLOTTED)
                            holeType = HoleTypeEnum.Slotted;
                        else if (boltGroup.HoleType == BoltGroup.BoltHoleTypeEnum.HOLE_TYPE_OVERSIZED)
                            holeType = HoleTypeEnum.Oversized;

                        slottedHoleX = boltGroup.SlottedHoleX.ToString();
                        slottedHoleY = boltGroup.SlottedHoleY.ToString();

                        if (boltGroup.RotateSlots == BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_ODD)
                            rotateSlots = RotateSlotEnum.Odd;
                        else if (boltGroup.RotateSlots == BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_EVEN)
                            rotateSlots = RotateSlotEnum.Even;
                        else if (boltGroup.RotateSlots == BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_PARALLEL)
                            rotateSlots = RotateSlotEnum.Parallel;

                        if (boltGroup.Position.Plane == Position.PlaneEnum.LEFT) positionOnPlane = PositionPlaneEnum.Left;
                        else if (boltGroup.Position.Plane == Position.PlaneEnum.MIDDLE) positionOnPlane = PositionPlaneEnum.Middle;
                        else if (boltGroup.Position.Plane == Position.PlaneEnum.RIGHT) positionOnPlane = PositionPlaneEnum.Right;
                        positionOnPlaneOffset = boltGroup.Position.PlaneOffset.ToString("F02");

                        if (boltGroup.Position.Rotation == Position.RotationEnum.FRONT) positionRotation = PositionRotationEnum.Front;
                        else if (boltGroup.Position.Rotation == Position.RotationEnum.TOP) positionRotation = PositionRotationEnum.Top;
                        else if (boltGroup.Position.Rotation == Position.RotationEnum.BACK) positionRotation = PositionRotationEnum.Back;
                        else if (boltGroup.Position.Rotation == Position.RotationEnum.BELOW) positionRotation = PositionRotationEnum.Below;
                        positionRotationOffset = boltGroup.Position.RotationOffset.ToString("F02");

                        if (boltGroup.Position.Depth == Position.DepthEnum.BEHIND) positionAtDepth = PositionDepthEnum.Behind;
                        else if (boltGroup.Position.Depth == Position.DepthEnum.FRONT) positionAtDepth = PositionDepthEnum.Front;
                        else if (boltGroup.Position.Depth == Position.DepthEnum.MIDDLE) positionAtDepth = PositionDepthEnum.Middle;
                        positionAtDepthOffset = boltGroup.Position.DepthOffset.ToString("F02");
                        offsetFromStartX = boltGroup.StartPointOffset.Dx.ToString("F02");
                        offsetFromStartY = boltGroup.StartPointOffset.Dy.ToString("F02");
                        offsetFromStartZ = boltGroup.StartPointOffset.Dz.ToString("F02");
                        offsetFromEndX = boltGroup.EndPointOffset.Dx.ToString("F02");
                        offsetFromEndY = boltGroup.EndPointOffset.Dy.ToString("F02");
                        offsetFromEndZ = boltGroup.EndPointOffset.Dz.ToString("F02");

                        if (boltGroup.Hole1) hole1 = Bool.True; else hole1 = Bool.False;
                        if (boltGroup.Hole2) hole2 = Bool.True; else hole2 = Bool.False;
                        if (boltGroup.Hole3) hole3 = Bool.True; else hole3 = Bool.False;
                        if (boltGroup.Hole4) hole4 = Bool.True; else hole4 = Bool.False;
                        if (boltGroup.Hole5) hole5 = Bool.True; else hole5 = Bool.False;

                        if (boltGroup.Washer1) washer1 = Bool.True; else washer1 = Bool.False;
                        if (boltGroup.Washer2) washer2 = Bool.True; else washer2 = Bool.False;
                        if (boltGroup.Washer3) washer3 = Bool.True; else washer3 = Bool.False;
                        
                        if (boltGroup.Nut1) nut1 = Bool.True; else nut1 = Bool.False;
                        if (boltGroup.Nut2) nut2 = Bool.True; else nut2 = Bool.False;

                        if (boltGroup.Bolt) boltShaft = Bool.True; else boltShaft = Bool.False;
                    }
                }
            }
            if (modelObjectEnum.GetSize() > 1)
            {
                boltSize = "";
                boltStandard = "";
                boltType = new BoltTypeEnum();
                threadMaterial = new ThreadMaterialEnum();
                cutLength = "";
                extraLength = "";
                boltGroupShape = "";
                boltDistX = "";
                boltDistY = "";
                tolerance = "";
                slottedHoleX = "";
                slottedHoleY = "";
                rotateSlots = new RotateSlotEnum();
                holeType = new HoleTypeEnum();
                positionOnPlane = new PositionPlaneEnum();
                positionOnPlaneOffset = "";
                positionRotation = new PositionRotationEnum();
                positionRotationOffset = "";
                positionAtDepth = new PositionDepthEnum();
                positionAtDepthOffset = "";
                offsetFromStartX = "";
                offsetFromStartY = "";
                offsetFromStartZ = "";
                offsetFromEndX = "";
                offsetFromEndY = "";
                offsetFromEndZ = "";
                hole1 = new Bool();
                hole2 = new Bool();
                hole3 = new Bool();
                hole4 = new Bool();
                hole5 = new Bool();
                boltShaft = new Bool();
                washer1 = new Bool();
                washer2 = new Bool();
                washer3 = new Bool();
                nut1 = new Bool();
                nut2 = new Bool();
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;

            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnum = modelObjectSelector.GetSelectedObjects();
            while (modelObjectEnum.MoveNext())
            {

                if (modelObjectEnum.Current is BoltGroup)
                {
                    BoltGroup bolt = (BoltGroup)modelObjectEnum.Current;

                    if (label == "BoltSize") bolt.BoltSize = double.Parse(boltSize);
                    if (label == "BoltStandard") bolt.BoltStandard = boltStandard; // list to sort
                    if (label == "BoltType")
                    {
                        if (boltType == BoltTypeEnum.Site) bolt.BoltType = BoltGroup.BoltTypeEnum.BOLT_TYPE_SITE;
                        if (boltType == BoltTypeEnum.Workshop) bolt.BoltType = BoltGroup.BoltTypeEnum.BOLT_TYPE_WORKSHOP;
                    }

                    if (label == "ThreadMaterial")
                    {
                        if (threadMaterial == ThreadMaterialEnum.No) bolt.ThreadInMaterial = BoltGroup.BoltThreadInMaterialEnum.THREAD_IN_MATERIAL_NO;
                        if (threadMaterial == ThreadMaterialEnum.Yes) bolt.ThreadInMaterial = BoltGroup.BoltThreadInMaterialEnum.THREAD_IN_MATERIAL_YES;
                    }

                    if (label == "CutLength") bolt.CutLength = double.Parse(cutLength);
                    if (label == "ExtraLength") bolt.ExtraLength = double.Parse(extraLength);
                    if (label == "BoltShaft") bolt.Bolt = bool.Parse(boltShaft.ToString());
                    if (label == "Washer1") bolt.Washer1 = bool.Parse(washer1.ToString());
                    if (label == "Washer2") bolt.Washer2 = bool.Parse(washer2.ToString());
                    if (label == "Washer3") bolt.Washer3 = bool.Parse(washer3.ToString());
                    if (label == "Nut1") bolt.Nut1 = bool.Parse(nut1.ToString());
                    if (label == "Nut2") bolt.Nut2 = bool.Parse(nut2.ToString());
                    
                    if (label == @"BoltDistX \ No Bolts")
                    {
                        if (bolt is BoltArray)
                        {
                            BoltArray boltArray = (BoltArray)bolt;
                            int boltArrayCount = boltArray.GetBoltDistXCount();
                            for (int i = boltArrayCount - 1; i >= 0; i--)
                            {
                                boltArray.RemoveBoltDistX(i);
                                boltArray.Modify();
                            }
                            if (boltDistX != "")
                            {
                                string[] splitSpaces = boltDistX.Split(new Char[] { ' ' });
                                System.Collections.ArrayList arr = new System.Collections.ArrayList();
                                foreach (string stringText in splitSpaces)
                                {
                                    if (stringText.Contains("*"))
                                    {
                                        string[] splitMultiplier = stringText.Split(new Char[] { '*' });
                                        for (int i = 0; i < int.Parse(splitMultiplier[0]); i++)
                                            arr.Add(double.Parse(splitMultiplier[1].ToString()));
                                    }
                                    else
                                        arr.Add(double.Parse(stringText));
                                }
                                for (int i = 0; i < arr.Count; i++)
                                    boltArray.SetBoltDistX(i, double.Parse(arr[i].ToString()));
                            }
                        }
                        else if (bolt is BoltXYList)
                            MessageBox.Show("Buggered");
                        else if (bolt is BoltCircle)
                        {
                            BoltCircle boltCircle = (BoltCircle)bolt;
                            boltCircle.NumberOfBolts = double.Parse(boltDistX);
                        }
                    }

                    if (label == @"BoltDistY \ Diameter")
                    {
                        if (bolt is BoltArray)
                        {
                            BoltArray boltArray = (BoltArray)bolt;
                            int boltArrayCount = boltArray.GetBoltDistYCount();
                            for (int i = boltArrayCount - 1; i >= 0; i--)
                            {
                                boltArray.RemoveBoltDistY(i);
                                boltArray.Modify();
                            }
                            if (boltDistY != "")
                            {
                                string[] splitSpaces = boltDistY.Split(new Char[] { ' ' });
                                System.Collections.ArrayList arr = new System.Collections.ArrayList();
                                foreach (string stringText in splitSpaces)
                                {
                                    if (stringText.Contains("*"))
                                    {
                                        string[] splitMultiplier = stringText.Split(new Char[] { '*' });
                                        for (int i = 0; i < int.Parse(splitMultiplier[0]); i++)
                                            arr.Add(double.Parse(splitMultiplier[1].ToString()));
                                    }
                                    else
                                        arr.Add(double.Parse(stringText));
                                }
                                for (int i = 0; i < arr.Count; i++)
                                    boltArray.SetBoltDistY(i, double.Parse(arr[i].ToString()));
                            }
                        }
                        else if (bolt is BoltXYList)
                        {
                            
                        }
                        else if (bolt is BoltCircle)
                        {
                            BoltCircle boltCircle = (BoltCircle)bolt;
                            boltCircle.Diameter = double.Parse(boltDistY);
                        }
                    }

                    if (label == "Tolerance") bolt.Tolerance = double.Parse(tolerance);
                    if (label == "HoleType")
                    {
                        if (holeType == HoleTypeEnum.Standard)
                        {
                            bolt.Hole1 = false;
                            bolt.Hole2 = false;
                            bolt.Hole3 = false;
                            bolt.Hole4 = false;
                            bolt.Hole5 = false;
                        }
                        if (holeType == HoleTypeEnum.Slotted) bolt.HoleType = BoltGroup.BoltHoleTypeEnum.HOLE_TYPE_SLOTTED;
                        if (holeType == HoleTypeEnum.Oversized) bolt.HoleType = BoltGroup.BoltHoleTypeEnum.HOLE_TYPE_OVERSIZED;
                    }

                    if (label == "SlottedHoleX") bolt.SlottedHoleX = double.Parse(slottedHoleX);
                    if (label == "SlottedHoleY") bolt.SlottedHoleY = double.Parse(slottedHoleY);
                    if (label == "RotateSlots")
                    {
                        if (rotateSlots == RotateSlotEnum.Even) bolt.RotateSlots = BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_EVEN;
                        if (rotateSlots == RotateSlotEnum.Odd) bolt.RotateSlots = BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_ODD;
                        if (rotateSlots == RotateSlotEnum.Parallel) bolt.RotateSlots = BoltGroup.BoltRotateSlotsEnum.ROTATE_SLOTS_PARALLEL;
                    }

                    if (label == "Hole1") bolt.Hole1 = bool.Parse(hole1.ToString());
                    if (label == "Hole2") bolt.Hole2 = bool.Parse(hole2.ToString());
                    if (label == "Hole3") bolt.Hole3 = bool.Parse(hole3.ToString());
                    if (label == "Hole4") bolt.Hole4 = bool.Parse(hole4.ToString());
                    if (label == "Hole5") bolt.Hole5 = bool.Parse(hole5.ToString());
                    if (label == "OffsetFromStartX") bolt.StartPointOffset.Dx = double.Parse(offsetFromStartX);
                    if (label == "OffsetFromStartY") bolt.StartPointOffset.Dy = double.Parse(offsetFromStartY);
                    if (label == "OffsetFromStartZ") bolt.StartPointOffset.Dz = double.Parse(offsetFromStartZ);
                    if (label == "OffsetFromEndX") bolt.EndPointOffset.Dx = double.Parse(offsetFromEndX);
                    if (label == "OffsetFromEndY") bolt.EndPointOffset.Dy = double.Parse(offsetFromEndY);
                    if (label == "OffsetFromEndZ") bolt.EndPointOffset.Dz = double.Parse(offsetFromEndZ);
                    if (label == "PositionOnPlaneOffset") bolt.Position.PlaneOffset = double.Parse(positionOnPlaneOffset);
                    if (label == "PositionRotation")
                    {
                        if (positionRotation == PositionRotationEnum.Back) bolt.Position.Rotation = Position.RotationEnum.BACK;
                        if (positionRotation == PositionRotationEnum.Below) bolt.Position.Rotation = Position.RotationEnum.BELOW;
                        if (positionRotation == PositionRotationEnum.Front) bolt.Position.Rotation = Position.RotationEnum.FRONT;
                        if (positionRotation == PositionRotationEnum.Top) bolt.Position.Rotation = Position.RotationEnum.TOP;
                    }

                    if (label == "PositionRotationOffset") bolt.Position.RotationOffset = double.Parse(positionRotationOffset);
                    if (label == "At Depth Offset") bolt.Position.DepthOffset = double.Parse(positionAtDepthOffset);

                    bolt.Modify();
                }
            }
        }
    }
}
