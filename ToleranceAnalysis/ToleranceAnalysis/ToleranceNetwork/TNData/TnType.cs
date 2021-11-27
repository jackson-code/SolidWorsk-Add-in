
namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    public enum TnGeometricFeatureType_e  // 節點種類
    {
        Edge = 0,
        Face = 1,
    }

    public enum TnGCType_e
    {
        // --------------------------- 結合 -------------------------- //
        MATE_TYPE = 0,


        // --------------------------- 幾何公差 -------------------------- //
        // 個別參考公差
        GTOL_STRAIGHT = 1, // 真直度
        GTOL_FLAT = 2,     // 真平度
        GTOL_CIRC = 3,     // 真圓度
        GTOL_CYL = 4,      // 圓柱度
        GTOL_LPROF = 5,    // 線輪廓度(直線輪廓)
        GTOL_SPROF = 6,    // 面輪廓度(曲面輪廓)

        // 交互參考公差
        GTOL_PARA = 7,      // 平行度
        GTOL_PERP = 8,      // 垂直度
        GTOL_ANGULAR = 9,   // 傾斜度
        GTOL_SRUN = 10,     // 圓偏轉度
        GTOL_TRUN = 11,     // 總偏轉度(全部偏轉)
        GTOL_POSI = 12,     // 位置度
        GTOL_CONC = 13,     // 同心度
                            // 對稱度


        // --------------------------- 尺寸公差 -------------------------- //
        // 個別參考公差
        PositionTol = 15,
        SizeTol = 16,
        AngularTol = 17,


        // --------------------------- Edge -------------------------- //
        Length = 18,
    }

    public enum TnMateType_e
    {
        MateCamTangent = 0,
        MateCoincident = 1,
        MateConcentric = 2,
        MateDistanceDim = 3,
        MateGearDim = 4,
        MateHinge = 5,

        MateInPlace = 6,
        MateLinearCoupler = 7,
        MateLock = 8,
        MateParallel = 9,
        MatePerpendicular = 10,

        MatePlanarAngleDim = 11,
        MateProfileCenter = 12,
        MateRackPinionDim = 13,
        MateScrew = 14,
        MateSlot = 15,

        MateSymmetric = 16,
        MateTangent = 17,
        MateUniversalJoint = 18,
        MateWidth = 19,
    }
}
