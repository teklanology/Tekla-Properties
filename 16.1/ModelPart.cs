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
    class ModelPart
    {
        Model model = new Model();
        
        private string partType;
        private string partID;
        private string owner;
        private string gridLocation;
        private string partPrefix;
        private string partStartNo;
        private string assemblyPrefix;
        private string assemblyStartNo;
        private string phase;
        private string name;
        private string profile;
        private string material;
        private string finish;
        private string classValue;
        private string userfield1;
        private string userfield2;
        private string userfield3;
        private string userfield4;
        private string notesComments;
        private string fittingNotes;
        private string fittingNotes2;
        private string cambering;
        private string paint;
        private string preliminaryMark;
        private string paintWFT;
        private string paintDFT;
        private string shearStart;
        private string shearEnd;
        private string axialStart;
        private string axialEnd;
        private string momentStart;
        private string momentEnd;
        private string connCodeStart;
        private string connCodeEnd;
        private PositionPlaneEnum positionOnPlane;
        private string positionOnPlaneOffset;
        private PositionRotationEnum positionRotation;
        private string positionRotationOffset;
        private PositionDepthEnum positionAtDepth;
        private string positionAtDepthOffset;

        [ReadOnly(true)]
        [Category("...")]
        public string PartType
        {
            get { return partType; }
            set { partType = value; }
        }

        [ReadOnly(true)]
        [Category("...")]
        public string PartID
        {
            get { return partID; }
            set { partID = value; }
        }

        [ReadOnly(true)]
        [Category("...")]
        public string Owner
        {
            get { return owner; }
            set { owner = value; }
        }

        [ReadOnly(true)]
        [Category("...")]
        public string GridLocation
        {
            get { return gridLocation; }
            set { gridLocation = value; }
        }

        [Category("Numbering Series")]
        public string PartPrefix
        {
            get { return partPrefix; }
            set { partPrefix = value; }
        }

        [Category("Numbering Series")]
        public string PartStartNo
        {
            get { return partStartNo; }
            set { partStartNo = value; }
        }

        [Category("Numbering Series")]
        public string AssemblyPrefix
        {
            get { return assemblyPrefix; }
            set { assemblyPrefix = value; }
        }

        [Category("Numbering Series")]
        public string AssemblyStartNo
        {
            get { return assemblyStartNo; }
            set { assemblyStartNo = value; }
        }

        [Category("...")]
        public string Phase
        {
            get { return phase; }
            set { phase = value; }
        }

        [Category("Attributes")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [Category("Attributes")]
        public string Profile
        {
            get { return profile; }
            set { profile = value; }
        }

        [Category("Attributes")]
        public string Material
        {
            get { return material; }
            set { material = value; }
        }

        [Category("Attributes")]
        public string Finish
        {
            get { return finish; }
            set { finish = value; }
        }

        [Category("Attributes")]
        public string Class
        {
            get { return classValue; }
            set { classValue = value; }
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
        public string NotesComments
        {
            get { return notesComments; }
            set { notesComments = value; }
        }

        [Category("User-defined Attributes")]
        public string FittingNotes
        {
            get { return fittingNotes; }
            set { fittingNotes = value; }
        }

        [Category("User-defined Attributes")]
        public string FittingNotes2
        {
            get { return fittingNotes2; }
            set { fittingNotes2 = value; }
        }

        [Category("User-defined Attributes")]
        public string Cambering
        {
            get { return cambering; }
            set { cambering = value; }
        }

        [Category("User-defined Attributes")]
        public string Paint
        {
            get { return paint; }
            set { paint = value; }
        }

        [Category("User-defined Attributes")]
        public string PreliminaryMark
        {
            get { return preliminaryMark; }
            set { preliminaryMark = value; }
        }

        [Category("User-defined Attributes")]
        public string PaintWFT
        {
            get { return paintWFT; }
            set { paintWFT = value; }
        }

        [Category("User-defined Attributes")]
        public string PaintDFT
        {
            get { return paintDFT; }
            set { paintDFT = value; }
        }

        [Category("User-defined Attributes")]
        public string ShearStart
        {
            get { return shearStart; }
            set { shearStart = value; }
        }

        [Category("User-defined Attributes")]
        public string ShearEnd
        {
            get { return shearEnd; }
            set { shearEnd = value; }
        }

        [Category("User-defined Attributes")]
        public string AxialStart
        {
            get { return axialStart; }
            set { axialStart = value; }
        }

        [Category("User-defined Attributes")]
        public string AxialEnd
        {
            get { return axialEnd; }
            set { axialEnd = value; }
        }

        [Category("User-defined Attributes")]
        public string MomentStart
        {
            get { return momentStart; }
            set { momentStart = value; }
        }

        [Category("User-defined Attributes")]
        public string MomentEnd
        {
            get { return momentEnd; }
            set { momentEnd = value; }
        }

        [Category("User-defined Attributes")]
        public string ConnCodeStart
        {
            get { return connCodeStart; }
            set { connCodeStart = value; }
        }

        [Category("User-defined Attributes")]
        public string ConnCodeEnd
        {
            get { return connCodeEnd; }
            set { connCodeEnd = value; }
        }

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
        
        public class Beam : ModelPart
        {
            private string zStart;
            private string zEnd;
            private string offsetStartPointX;
            private string offsetStartPointY;
            private string offsetStartPointZ;
            private string offsetEndPointX;
            private string offsetEndPointY;
            private string offsetEndPointZ;

            public string ZStart
            {
                get { return zStart; }
                set { zStart = value; }
            }

            public string ZEnd
            {
                get { return zEnd; }
                set { zEnd = value; }
            }

            public string OffsetStartPointX
            {
                get { return offsetStartPointX; }
                set { offsetStartPointX = value; }
            }

            public string OffsetStartPointY
            {
                get { return offsetStartPointY; }
                set { offsetStartPointY = value; }
            }

            public string OffsetStartPointZ
            {
                get { return offsetStartPointZ; }
                set { offsetStartPointZ = value; }
            }

            public string OffsetEndPointX
            {
                get { return offsetEndPointX; }
                set { offsetEndPointX = value; }
            }

            public string OffsetEndPointY
            {
                get { return offsetEndPointY; }
                set { offsetEndPointY = value; }
            }

            public string OffsetEndPointZ
            {
                get { return offsetEndPointZ; }
                set { offsetEndPointZ = value; }
            }

            public void GetProperties()
            {
                ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
                if (modelObjectEnum.GetSize() == 1)
                {
                    while (modelObjectEnum.MoveNext())
                    {
                        if (modelObjectEnum.Current is Tekla.Structures.Model.Beam)
                        {
                            Tekla.Structures.Model.Beam beam = (Tekla.Structures.Model.Beam)modelObjectEnum.Current;

                            TransformationPlane currentTP = new TransformationPlane();
                            currentTP = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();

                            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());

                            beam.Select();

                            Assembly assembly = beam.GetAssembly() as Tekla.Structures.Model.Assembly;
                            assembly.GetReportProperty("ASSEMBLY_POSITION_CODE", ref gridLocation);
                            partType = beam.GetType().Name;
                            partID = beam.Identifier.ID.ToString();
                            beam.GetReportProperty("OWNER", ref owner);
                            zStart = beam.StartPoint.Z.ToString("F02");
                            zEnd = beam.EndPoint.Z.ToString("F02");
                            partPrefix = beam.PartNumber.Prefix;
                            partStartNo = beam.PartNumber.StartNumber.ToString();
                            assemblyPrefix = beam.AssemblyNumber.Prefix;
                            assemblyStartNo = beam.AssemblyNumber.StartNumber.ToString();
                            Phase CurrentPhase = new Phase();
                            beam.GetPhase(out CurrentPhase);
                            phase = CurrentPhase.PhaseNumber.ToString();
                            name = beam.Name;
                            profile = beam.Profile.ProfileString;
                            material = beam.Material.MaterialString;
                            finish = beam.Finish;
                            classValue = beam.Class;
                            if (beam.Position.Plane == Position.PlaneEnum.LEFT) positionOnPlane = PositionPlaneEnum.Left;
                            else if (beam.Position.Plane == Position.PlaneEnum.MIDDLE) positionOnPlane = PositionPlaneEnum.Middle;
                            else if (beam.Position.Plane == Position.PlaneEnum.RIGHT) positionOnPlane = PositionPlaneEnum.Right;
                            positionOnPlaneOffset = beam.Position.PlaneOffset.ToString("F02");
                            if (beam.Position.Rotation == Position.RotationEnum.BACK) positionRotation = PositionRotationEnum.Back;
                            else if (beam.Position.Rotation == Position.RotationEnum.BELOW) positionRotation = PositionRotationEnum.Below;
                            else if (beam.Position.Rotation == Position.RotationEnum.FRONT) positionRotation = PositionRotationEnum.Front;
                            else if (beam.Position.Rotation == Position.RotationEnum.TOP) positionRotation = PositionRotationEnum.Top;
                            positionRotationOffset = beam.Position.RotationOffset.ToString("F02");
                            if (beam.Position.Depth == Position.DepthEnum.BEHIND) positionAtDepth = PositionDepthEnum.Behind;
                            else if (beam.Position.Depth == Position.DepthEnum.FRONT) positionAtDepth = PositionDepthEnum.Front;
                            else if (beam.Position.Depth == Position.DepthEnum.MIDDLE) positionAtDepth = PositionDepthEnum.Middle;
                            positionAtDepthOffset = beam.Position.DepthOffset.ToString("F02");
                            offsetStartPointX = beam.StartPointOffset.Dx.ToString("F02");
                            offsetStartPointY = beam.StartPointOffset.Dy.ToString("F02");
                            offsetStartPointZ = beam.StartPointOffset.Dz.ToString("F02");
                            offsetEndPointX = beam.EndPointOffset.Dx.ToString("F02");
                            offsetEndPointY = beam.EndPointOffset.Dy.ToString("F02");
                            offsetEndPointZ = beam.EndPointOffset.Dz.ToString("F02");
                            beam.GetUserProperty("USER_FIELD_1", ref userfield1);
                            beam.GetUserProperty("USER_FIELD_2", ref userfield2);
                            beam.GetUserProperty("USER_FIELD_3", ref userfield3);
                            beam.GetUserProperty("USER_FIELD_4", ref userfield4);
                            beam.GetUserProperty("comment", ref notesComments);
                            beam.GetUserProperty("FIT_NOTES", ref fittingNotes);
                            beam.GetUserProperty("FIT_NOTES2", ref fittingNotes2);
                            beam.GetUserProperty("cambering", ref cambering);
                            beam.GetUserProperty("PAINT", ref paint);
                            beam.GetUserProperty("PRELIM_MARK", ref preliminaryMark);
                            beam.GetUserProperty("PAINT_WFT", ref paintWFT);
                            beam.GetUserProperty("PAINT_DFT", ref paintDFT);
                            double dblshear1 = 0; beam.GetUserProperty("shear1", ref dblshear1); dblshear1 = dblshear1 * 0.001; shearStart = dblshear1.ToString();
                            double dblshear2 = 0; beam.GetUserProperty("shear2", ref dblshear2); dblshear2 = dblshear2 * 0.001; shearEnd = dblshear2.ToString();
                            double dblaxial1 = 0; beam.GetUserProperty("axial1", ref dblaxial1); dblaxial1 = dblaxial1 * 0.001; axialStart = dblaxial1.ToString();
                            double dblaxial2 = 0; beam.GetUserProperty("axial2", ref dblaxial2); dblaxial2 = dblaxial2 * 0.001; axialEnd = dblaxial2.ToString();
                            double dblmoment1 = 0; beam.GetUserProperty("moment1", ref dblmoment1); dblmoment1 = dblmoment1 * 0.001; momentStart = dblmoment1.ToString();
                            double dblmoment2 = 0; beam.GetUserProperty("moment2", ref dblmoment2); dblmoment2 = dblmoment2 * 0.001; momentEnd = dblmoment2.ToString();

                            beam.GetUserProperty("CONN_CODE_END1", ref connCodeStart);
                            beam.GetUserProperty("CONN_CODE_END2", ref connCodeEnd);
                            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentTP);
                        }
                    }
                }
                if (modelObjectEnum.GetSize() > 1)
                {
                    partType = "";
                    partID = "";
                    owner = "";
                    gridLocation = "";
                    partPrefix = "";
                    partStartNo = "";
                    assemblyPrefix = "";
                    assemblyStartNo = "";
                    phase = "";
                    name = "";
                    profile = "";
                    material = "";
                    finish = "";
                    classValue = "";
                    userfield1 = "";
                    userfield2 = "";
                    userfield3 = "";
                    userfield4 = "";
                    notesComments = "";
                    fittingNotes = "";
                    fittingNotes2 = "";
                    cambering = "";
                    paint = "";
                    preliminaryMark = "";
                    paintWFT = "";
                    paintDFT = "";
                    shearStart = "";
                    shearEnd = "";
                    axialStart = "";
                    axialEnd = "";
                    momentStart = "";
                    momentEnd = "";
                    connCodeStart = "";
                    connCodeEnd = "";
                    positionOnPlane = new PositionPlaneEnum();
                    positionOnPlaneOffset = "";
                    positionRotation = new PositionRotationEnum();
                    positionRotationOffset = "";
                    positionAtDepth = new PositionDepthEnum();
                    positionAtDepthOffset = "";
                    offsetStartPointX = "";
                    offsetStartPointY = "";
                    offsetStartPointZ = "";
                    offsetEndPointX = "";
                    offsetEndPointY = "";
                    offsetEndPointZ = "";
                    zStart = "";
                    zEnd = "";
                }
            }

            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
                while (modelObjectEnum.MoveNext())
                {
                    if (modelObjectEnum.Current is Tekla.Structures.Model.Beam)
                    {
                        Tekla.Structures.Model.Beam beam = (Tekla.Structures.Model.Beam)modelObjectEnum.Current;
                        if (label == "Name") beam.Name = name;
                        if (label == "Profile") beam.Profile.ProfileString = profile;
                        if (label == "Material") beam.Material.MaterialString = material;
                        if (label == "Finish") beam.Finish = finish;
                        if (label == "Class") beam.Class = classValue;
                        if (label == "PartPrefix") beam.PartNumber.Prefix = partPrefix;
                        if (label == "PartStartNo") beam.PartNumber.StartNumber = int.Parse(partStartNo);
                        if (label == "AssemblyPrefix") beam.AssemblyNumber.Prefix = assemblyPrefix;
                        if (label == "AssemblyStartNo") beam.AssemblyNumber.StartNumber = int.Parse(assemblyStartNo);
                        if (label == "Phase") beam.SetPhase(new Phase(int.Parse(phase)));
                        if (label == "OffsetStartPointX") beam.StartPointOffset.Dx = double.Parse(offsetStartPointX);
                        if (label == "OffsetStartPointY") beam.StartPointOffset.Dy = double.Parse(offsetStartPointY);
                        if (label == "OffsetStartPointZ") beam.StartPointOffset.Dz = double.Parse(offsetStartPointZ);
                        if (label == "OffsetEndPointX") beam.EndPointOffset.Dx = double.Parse(offsetEndPointX);
                        if (label == "OffsetEndPointY") beam.EndPointOffset.Dy = double.Parse(offsetEndPointY);
                        if (label == "OffsetEndPointZ") beam.EndPointOffset.Dz = double.Parse(offsetEndPointZ);
                        if (label == "ZStart") beam.StartPoint.Z = double.Parse(zStart);
                        if (label == "ZEnd") beam.EndPoint.Z = double.Parse(zEnd);
                        if (label == "PositionOnPlane")
                        {
                            if (positionOnPlane == PositionPlaneEnum.Left) beam.Position.Plane = Position.PlaneEnum.LEFT;
                            if (positionOnPlane == PositionPlaneEnum.Middle) beam.Position.Plane = Position.PlaneEnum.MIDDLE;
                            if (positionOnPlane == PositionPlaneEnum.Right) beam.Position.Plane = Position.PlaneEnum.RIGHT;
                        }

                        if (label == "PositionOnPlaneOffset") beam.Position.PlaneOffset = double.Parse(positionOnPlaneOffset);
                        
                        if (label == "PositionRotation")
                        {
                            if (positionRotation == PositionRotationEnum.Top) beam.Position.Rotation = Position.RotationEnum.TOP;
                            if (positionRotation == PositionRotationEnum.Front) beam.Position.Rotation = Position.RotationEnum.FRONT;
                            if (positionRotation == PositionRotationEnum.Back) beam.Position.Rotation = Position.RotationEnum.BACK;
                            if (positionRotation == PositionRotationEnum.Below) beam.Position.Rotation = Position.RotationEnum.BELOW;
                        }

                        if (label == "PositionRotationOffset") beam.Position.RotationOffset = double.Parse(positionRotationOffset);

                        if (label == "PositionAtDepth")
                        {
                            if (positionAtDepth == PositionDepthEnum.Behind) beam.Position.Depth = Position.DepthEnum.BEHIND;
                            if (positionAtDepth == PositionDepthEnum.Front) beam.Position.Depth = Position.DepthEnum.FRONT;
                            if (positionAtDepth == PositionDepthEnum.Middle) beam.Position.Depth = Position.DepthEnum.MIDDLE;
                        }

                        if (label == "PositionAtDepthOffset") beam.Position.DepthOffset = double.Parse(positionAtDepthOffset);
                        if (label == "Userfield1") beam.SetUserProperty("USER_FIELD_1", userfield1);
                        if (label == "Userfield2") beam.SetUserProperty("USER_FIELD_2", userfield2);
                        if (label == "Userfield3") beam.SetUserProperty("USER_FIELD_3", userfield3);
                        if (label == "Userfield4") beam.SetUserProperty("USER_FIELD_4", userfield4);
                        if (label == "NotesComments") beam.SetUserProperty("comment", notesComments);
                        if (label == "FittingNotes") beam.SetUserProperty("FIT_NOTES", fittingNotes);
                        if (label == "FittingNotes2") beam.SetUserProperty("FIT_NOTES2", fittingNotes2);
                        if (label == "Cambering") beam.SetUserProperty("cambering", cambering);
                        if (label == "Paint") beam.SetUserProperty("PAINT", paint);
                        if (label == "PreliminaryMark") beam.SetUserProperty("PRELIM_MARK", preliminaryMark);
                        if (label == "PaintWFT") beam.SetUserProperty("PAINT_WFT", paintWFT);
                        if (label == "PaintDFT") beam.SetUserProperty("PAINT_DFT", paintDFT);
                        if (label == "ShearStart") beam.SetUserProperty("shear1", double.Parse(shearStart) * 1000);
                        if (label == "ShearEnd") beam.SetUserProperty("shear2", double.Parse(shearEnd) * 1000);
                        if (label == "AxialStart") beam.SetUserProperty("axial1", double.Parse(axialStart) * 1000);
                        if (label == "AxialEnd") beam.SetUserProperty("axial2", double.Parse(axialEnd) * 1000);
                        if (label == "MomentStart") beam.SetUserProperty("moment1", double.Parse(momentStart) * 1000);
                        if (label == "MomentEnd") beam.SetUserProperty("moment2", double.Parse(momentEnd) * 1000);
                        if (label == "ConnCodeStart") beam.SetUserProperty("CONN_CODE_END1", connCodeStart);
                        if (label == "ConnCodeEnd") beam.SetUserProperty("CONN_CODE_END2", connCodeEnd);

                        beam.Modify();
                    }
                }
            }
        }

        public class ContourPlate : ModelPart
        {
            public void GetProperties()
            {
                ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
                if (modelObjectEnum.GetSize() == 1)
                {
                    while (modelObjectEnum.MoveNext())
                    {
                        if (modelObjectEnum.Current is Tekla.Structures.Model.ContourPlate)
                        {
                            Tekla.Structures.Model.ContourPlate contourPlate = (Tekla.Structures.Model.ContourPlate)modelObjectEnum.Current;

                            TransformationPlane currentTP = new TransformationPlane();
                            currentTP = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();

                            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());

                            contourPlate.Select();

                            Assembly assembly = contourPlate.GetAssembly() as Tekla.Structures.Model.Assembly;
                            assembly.GetReportProperty("ASSEMBLY_POSITION_CODE", ref gridLocation);
                            partType = contourPlate.GetType().Name;
                            partID = contourPlate.Identifier.ID.ToString();
                            contourPlate.GetReportProperty("OWNER", ref owner);
                            partPrefix = contourPlate.PartNumber.Prefix;
                            partStartNo = contourPlate.PartNumber.StartNumber.ToString();
                            assemblyPrefix = contourPlate.AssemblyNumber.Prefix;
                            assemblyStartNo = contourPlate.AssemblyNumber.StartNumber.ToString();
                            Phase CurrentPhase = new Phase();
                            contourPlate.GetPhase(out CurrentPhase);
                            phase = CurrentPhase.PhaseNumber.ToString();
                            name = contourPlate.Name;
                            profile = contourPlate.Profile.ProfileString;
                            material = contourPlate.Material.MaterialString;
                            finish = contourPlate.Finish;
                            classValue = contourPlate.Class;

                            if (contourPlate.Position.Plane == Position.PlaneEnum.LEFT) positionOnPlane = PositionPlaneEnum.Left;
                            else if (contourPlate.Position.Plane == Position.PlaneEnum.MIDDLE) positionOnPlane = PositionPlaneEnum.Middle;
                            else if (contourPlate.Position.Plane == Position.PlaneEnum.RIGHT) positionOnPlane = PositionPlaneEnum.Right;

                            positionOnPlaneOffset = contourPlate.Position.PlaneOffset.ToString("F02");

                            if (contourPlate.Position.Rotation == Position.RotationEnum.BACK) positionRotation = PositionRotationEnum.Back;
                            else if (contourPlate.Position.Rotation == Position.RotationEnum.BELOW) positionRotation = PositionRotationEnum.Below;
                            else if (contourPlate.Position.Rotation == Position.RotationEnum.FRONT) positionRotation = PositionRotationEnum.Front;
                            else if (contourPlate.Position.Rotation == Position.RotationEnum.TOP) positionRotation = PositionRotationEnum.Top;

                            positionRotationOffset = contourPlate.Position.RotationOffset.ToString("F02");

                            if (contourPlate.Position.Depth == Position.DepthEnum.BEHIND) positionAtDepth = PositionDepthEnum.Behind;
                            else if (contourPlate.Position.Depth == Position.DepthEnum.FRONT) positionAtDepth = PositionDepthEnum.Front;
                            else if (contourPlate.Position.Depth == Position.DepthEnum.MIDDLE) positionAtDepth = PositionDepthEnum.Middle;

                            positionAtDepthOffset = contourPlate.Position.DepthOffset.ToString("F02");
                            contourPlate.GetUserProperty("USER_FIELD_1", ref userfield1);
                            contourPlate.GetUserProperty("USER_FIELD_2", ref userfield2);
                            contourPlate.GetUserProperty("USER_FIELD_3", ref userfield3);
                            contourPlate.GetUserProperty("USER_FIELD_4", ref userfield4);
                            contourPlate.GetUserProperty("comment", ref notesComments);
                            contourPlate.GetUserProperty("FIT_NOTES", ref fittingNotes);
                            contourPlate.GetUserProperty("FIT_NOTES2", ref fittingNotes2);
                            contourPlate.GetUserProperty("cambering", ref cambering);
                            contourPlate.GetUserProperty("PAINT", ref paint);
                            contourPlate.GetUserProperty("PRELIM_MARK", ref preliminaryMark);
                            contourPlate.GetUserProperty("PAINT_WFT", ref paintWFT);
                            contourPlate.GetUserProperty("PAINT_DFT", ref paintDFT);
                            double dblshear1 = 0; contourPlate.GetUserProperty("shear1", ref dblshear1); dblshear1 = dblshear1 * 0.001; shearStart = dblshear1.ToString();
                            double dblshear2 = 0; contourPlate.GetUserProperty("shear2", ref dblshear2); dblshear2 = dblshear2 * 0.001; shearEnd = dblshear2.ToString();
                            double dblaxial1 = 0; contourPlate.GetUserProperty("axial1", ref dblaxial1); dblaxial1 = dblaxial1 * 0.001; axialStart = dblaxial1.ToString();
                            double dblaxial2 = 0; contourPlate.GetUserProperty("axial2", ref dblaxial2); dblaxial2 = dblaxial2 * 0.001; axialEnd = dblaxial2.ToString();
                            double dblmoment1 = 0; contourPlate.GetUserProperty("moment1", ref dblmoment1); dblmoment1 = dblmoment1 * 0.001; momentStart = dblmoment1.ToString();
                            double dblmoment2 = 0; contourPlate.GetUserProperty("moment2", ref dblmoment2); dblmoment2 = dblmoment2 * 0.001; momentEnd = dblmoment2.ToString();

                            contourPlate.GetUserProperty("CONN_CODE_END1", ref connCodeStart);
                            contourPlate.GetUserProperty("CONN_CODE_END2", ref connCodeEnd);
                            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentTP);
                        }
                    }
                }
                if (modelObjectEnum.GetSize() > 1)
                {
                    partType = "";
                    partID = "";
                    owner = "";
                    gridLocation = "";
                    partPrefix = "";
                    partStartNo = "";
                    assemblyPrefix = "";
                    assemblyStartNo = "";
                    phase = "";
                    name = "";
                    profile = "";
                    material = "";
                    finish = "";
                    classValue = "";
                    userfield1 = "";
                    userfield2 = "";
                    userfield3 = "";
                    userfield4 = "";
                    notesComments = "";
                    fittingNotes = "";
                    fittingNotes2 = "";
                    cambering = "";
                    paint = "";
                    preliminaryMark = "";
                    paintWFT = "";
                    paintDFT = "";
                    shearStart = "";
                    shearEnd = "";
                    axialStart = "";
                    axialEnd = "";
                    momentStart = "";
                    momentEnd = "";
                    connCodeStart = "";
                    connCodeEnd = "";
                    positionOnPlane = new PositionPlaneEnum();
                    positionOnPlaneOffset = "";
                    positionRotation = new PositionRotationEnum();
                    positionRotationOffset = "";
                    positionAtDepth = new PositionDepthEnum();
                    positionAtDepthOffset = "";
                }
            }
            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                ModelObjectEnumerator modelObjectEnum = model.GetModelObjectSelector().GetSelectedObjects();
                while (modelObjectEnum.MoveNext())
                {
                    if (modelObjectEnum.Current is Tekla.Structures.Model.ContourPlate)
                    {
                        Tekla.Structures.Model.ContourPlate contourPlate = (Tekla.Structures.Model.ContourPlate)modelObjectEnum.Current;
                        if (label == "Name") contourPlate.Name = name;
                        if (label == "Profile") contourPlate.Profile.ProfileString = profile;
                        if (label == "Material") contourPlate.Material.MaterialString = material;
                        if (label == "Finish") contourPlate.Finish = finish;
                        if (label == "Class") contourPlate.Class = classValue;
                        if (label == "PartPrefix") contourPlate.PartNumber.Prefix = partPrefix;
                        if (label == "PartStartNo") contourPlate.PartNumber.StartNumber = int.Parse(partStartNo);
                        if (label == "AssemblyPrefix") contourPlate.AssemblyNumber.Prefix = assemblyPrefix;
                        if (label == "AssemblyStartNo") contourPlate.AssemblyNumber.StartNumber = int.Parse(assemblyStartNo);
                        if (label == "Phase") contourPlate.SetPhase(new Phase(int.Parse(phase)));
                        
                        if (label == "PositionOnPlane")
                        {
                            if (positionOnPlane == PositionPlaneEnum.Left) contourPlate.Position.Plane = Position.PlaneEnum.LEFT;
                            if (positionOnPlane == PositionPlaneEnum.Middle) contourPlate.Position.Plane = Position.PlaneEnum.MIDDLE;
                            if (positionOnPlane == PositionPlaneEnum.Right) contourPlate.Position.Plane = Position.PlaneEnum.RIGHT;
                        }

                        if (label == "PositionOnPlaneOffset") contourPlate.Position.PlaneOffset = double.Parse(positionOnPlaneOffset);
                        
                        if (label == "PositionRotation")
                        {
                            if (positionRotation == PositionRotationEnum.Top) contourPlate.Position.Rotation = Position.RotationEnum.TOP;
                            if (positionRotation == PositionRotationEnum.Front) contourPlate.Position.Rotation = Position.RotationEnum.FRONT;
                            if (positionRotation == PositionRotationEnum.Back) contourPlate.Position.Rotation = Position.RotationEnum.BACK;
                            if (positionRotation == PositionRotationEnum.Below) contourPlate.Position.Rotation = Position.RotationEnum.BELOW;
                        }

                        if (label == "PositionRotationOffset") contourPlate.Position.RotationOffset = double.Parse(positionRotationOffset);

                        if (label == "PositionAtDepth")
                        {
                            if (positionAtDepth == PositionDepthEnum.Behind) contourPlate.Position.Depth = Position.DepthEnum.BEHIND;
                            if (positionAtDepth == PositionDepthEnum.Front) contourPlate.Position.Depth = Position.DepthEnum.FRONT;
                            if (positionAtDepth == PositionDepthEnum.Middle) contourPlate.Position.Depth = Position.DepthEnum.MIDDLE;
                        }
                        if (label == "AtDepthOffset") contourPlate.Position.DepthOffset = double.Parse(positionAtDepthOffset);
                        if (label == "Userfield1") contourPlate.SetUserProperty("USER_FIELD_1", userfield1);
                        if (label == "Userfield2") contourPlate.SetUserProperty("USER_FIELD_2", userfield2);
                        if (label == "Userfield3") contourPlate.SetUserProperty("USER_FIELD_3", userfield3);
                        if (label == "Userfield4") contourPlate.SetUserProperty("USER_FIELD_4", userfield4);
                        if (label == "NotesComments") contourPlate.SetUserProperty("comment", notesComments);
                        if (label == "FittingNotes") contourPlate.SetUserProperty("FIT_NOTES", fittingNotes);
                        if (label == "FittingNotes2") contourPlate.SetUserProperty("FIT_NOTES2", fittingNotes2);
                        if (label == "Cambering") contourPlate.SetUserProperty("cambering", cambering);
                        if (label == "Paint") contourPlate.SetUserProperty("PAINT", paint);
                        if (label == "PreliminaryMark") contourPlate.SetUserProperty("PRELIM_MARK", preliminaryMark);
                        if (label == "PaintWFT") contourPlate.SetUserProperty("PAINT_WFT", paintWFT);
                        if (label == "PaintDFT") contourPlate.SetUserProperty("PAINT_DFT", paintDFT);
                        if (label == "ShearStart") contourPlate.SetUserProperty("shear1", double.Parse(shearStart) * 1000);
                        if (label == "ShearEnd") contourPlate.SetUserProperty("shear2", double.Parse(shearEnd) * 1000);
                        if (label == "AxialStart") contourPlate.SetUserProperty("axial1", double.Parse(axialStart) * 1000);
                        if (label == "AxialEnd") contourPlate.SetUserProperty("axial2", double.Parse(axialEnd) * 1000);
                        if (label == "MomentStart") contourPlate.SetUserProperty("moment1", double.Parse(momentStart) * 1000);
                        if (label == "MomentEnd") contourPlate.SetUserProperty("moment2", double.Parse(momentEnd) * 1000);
                        if (label == "ConnCodeStart") contourPlate.SetUserProperty("CONN_CODE_END1", connCodeStart);
                        if (label == "ConnCodeEnd") contourPlate.SetUserProperty("CONN_CODE_END2", connCodeEnd);

                        contourPlate.Modify();
                        model.CommitChanges();
                    }
                }
            }
        }
    }
}
