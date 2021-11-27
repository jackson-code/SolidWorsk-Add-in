using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.Construct
{
    static class SurfaceProcessor
    {
        public static void GetInfo(Face2 swFace, ref string surfaceType, ref double[] surfaceParam)
        {
            Surface swSurf = swFace.GetSurface();

            if (swSurf.IsBlending())
            {

            }
            else if (swSurf.IsCone())
            {
                surfaceType = "Cone";
                surfaceParam = swSurf.ConeParams2;
            }
            else if (swSurf.IsCylinder())
            {
                surfaceType = "Cylinder";
                surfaceParam = swSurf.CylinderParams;
            }
            else if (swSurf.IsForeign())
            {

            }
            else if (swSurf.IsOffset())
            {

            }
            else if (swSurf.IsParametric())
            {

            }
            else if (swSurf.IsPlane())
            {
                surfaceType = "Plane";
                surfaceParam = swSurf.PlaneParams;
            }
            else if (swSurf.IsRevolved())
            {

            }
            else if (swSurf.IsSphere())
            {
                surfaceType = "Sphere";
                surfaceParam = swSurf.SphereParams;
            }
            else if (swSurf.IsSwept())
            {

            }
            else if (swSurf.IsTorus())
            {
                surfaceType = "Torus";
                surfaceParam = swSurf.TorusParams;
            }
            else
            {
                surfaceType = "Unknown Surface Type";
                surfaceParam = null;
            }
        }
    }
}
