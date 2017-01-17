using System;

namespace TeklaProperties
{
    public enum ArrowheadPositionEnum 
    { 
        None = 1, 
        Start = 2, 
        End = 3, 
        Both = 4,
    }
    
    public enum ArrowheadTypeEnum 
    { 
        FilledArrow = 1, 
        LineArrow = 2, 
        CircleArrow = 3, 
        FilledCircleArrow = 4, 
    }

    public enum BoltTypeEnum 
    { 
        Site = 1, 
        Workshop = 2, 
    }

    public enum Bool 
    { 
        True = 1, 
        False = 2, 
    }

    public enum ContourType 
    { 
        None = 1, 
        Flush = 2,
    }

    public enum DimensionUnitsEnum 
    { 
        Automatic = 1, 
        Millimeter = 2, 
        Centimeter = 3, 
        Meter = 4, 
        Inch = 5, 
    }

    public enum FontNameEnum
    {
        Arial = 1,
        Comic_Sans_MS = 2,
    }

    public enum FrameTypeEnum
    {
        None = 1, 
        Rectangular = 2, 
        Line = 3, 
        Round = 4, 
        Circle = 5,
        Diamond = 6, 
        Hexagon = 7, 
        Triangle = 8, 
        Sharpened = 9,
    }

    public enum HoleTypeEnum 
    { 
        Standard = 1, 
        Slotted = 2, 
        Oversized = 3,
    }

    public enum LineTypeEnum
    {
        UndefinedLine = 1,
        SolidLine = 2,
        DashedLine = 3,
        SlashedLine = 4,
        DashDot = 5,
        DottedLine = 6,
        DashDoubleDot = 7,
        SlashDash = 8,
    }

    public enum MaterialTypeEnum
    {
        S275_I = 1,
        S275_E = 2,
        S275JR = 3,
        S275JO = 4,
        S355_I = 5,
        S355_E = 6,
        S355JR = 7,
        S355JO = 8,
        S355J2 = 9,
        S355J2G4 = 10,
        S355K2 = 11,
        S355K2G3 = 12,
    }

    public enum PlacingEnum 
    { 
        Free = 1, 
        Fixed = 2, 
    }

    public enum PlacingTypeEnum
    {
        AlongLinePlacing = 1,
        BaseLinePlacing = 2,
        LeaderLinePlacing = 3,
        PointPlacing = 4,
    }


    public enum PositionDepthEnum
    {
        Middle = 1,
        Front = 2,
        Behind = 3,
    }

    public enum PositionPlaneEnum
    {
        Middle = 1,
        Left = 2,
        Right = 3,
    }

    public enum PositionRotationEnum 
    { 
        Front = 1, 
        Top = 2, 
        Back = 3, 
        Below = 4,
    }

    public enum PreferredPlacingEnum 
    {
        PointPlacingType = 1,
        LeaderLinePlacingType = 2, 
    }

    public enum RotateSlotEnum 
    { 
        Odd = 1, 
        Even = 2, 
        Parallel = 3, 
    }

    public enum ShortDimensionEnum 
    { 
        Inside = 1, 
        Outside = 2,
    }
    
    public enum TextAlignmentEnum 
    { 
        Left = 1, 
        Center = 2, 
        Right = 3,
    }

    public enum ThreadMaterialEnum 
    { 
        Yes = 1, 
        No = 2,
    }
    
    public enum WeldType
    {
        WELD_TYPE_NONE, 
        WELD_TYPE_FILLET, 
        WELD_TYPE_BEVEL_GROOVE_SINGLE_V_BUTT, 
        WELD_TYPE_BEVEL_GROOVE_SINGLE_BEVEL_BUTT,
        WELD_TYPE_SQUARE_GROOVE_SQUARE_BUTT, 
        WELD_TYPE_SINGLE_V_BUTT_WITH_BROAD_ROOT_FACE, 
        WELD_TYPE_SINGLE_BEVEL_BUTT_WITH_BROAD_ROOT_FACE,
        WELD_TYPE_U_GROOVE_SINGLE_U_BUTT, 
        WELD_TYPE_J_GROOVE_J_BUTT, 
        WELD_TYPE_FLARE_V_GROOVE, 
        WELD_TYPE_FLARE_BEVEL_GROOVE,
        WELD_TYPE_EDGE_FLANGE, 
        WELD_TYPE_CORNER_FLANGE, 
        WELD_TYPE_PLUG, 
        WELD_TYPE_BEVEL_BACKING, 
        WELD_TYPE_SPOT, 
        WELD_TYPE_SEAM,
        WELD_TYPE_SLOT, 
        WELD_TYPE_PARTIAL_PENETRATION_SINGLE_BEVEL_BUTT_PLUS_FILLET, 
        WELD_TYPE_PARTIAL_PENETRATION_SQUARE_GROOVE_PLUS_FILLET,
        WELD_TYPE_MELT_THROUGH, 
        STEEP_FLANKED_BEVEL_GROOVE_SINGLE_V_BUTT, 
        STEEP_FLANKED_BEVEL_GROOVE_SINGLE_BEVEL_BUTT,
        WELD_TYPE_EDGE, 
        WELD_TYPE_ISO_SURFACING, 
        WELD_TYPE_FOLD, 
        WELD_TYPE_INCLINED,
    }
    
}
