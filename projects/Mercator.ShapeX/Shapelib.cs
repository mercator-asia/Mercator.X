using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Mercator.ShapeX
{
    /// <summary>
    /// .NET Framework wrapper for Shapefile C Library V1.2.10
    /// </summary>
    /// <remarks>
    /// Shapefile C Library is (c) 1998 Frank Warmerdam.  .NET wrapper provided by David Gancarz.  
    /// Please send error reports or other suggestions regarding this wrapper class to:
    /// dgancarz@cfl.rr.com or david.gancarz@cityoforlando.net
    /// </remarks>

    public class Shapelib
    {
        /// <summary>
        /// Shape type enumeration
        /// </summary>
        public enum ShapeType
        {
            /// <summary>Shape with no geometric data</summary>
            NullShape = 0,
            /// <summary>2D point</summary>
            Point = 1,
            /// <summary>2D polyline</summary>
            PolyLine = 3,
            /// <summary>2D polygon</summary>
            Polygon = 5,
            /// <summary>Set of 2D points</summary>
            MultiPoint = 8,
            /// <summary>3D point</summary>
            PointZ = 11,
            /// <summary>3D polyline</summary>
            PolyLineZ = 13,
            /// <summary>3D polygon</summary>
            PolygonZ = 15,
            /// <summary>Set of 3D points</summary>
            MultiPointZ = 18,
            /// <summary>3D point with measure</summary>
            PointM = 21,
            /// <summary>3D polyline with measure</summary>
            PolyLineM = 23,
            /// <summary>3D polygon with measure</summary>
            PolygonM = 25,
            /// <summary>Set of 3d points with measures</summary>
            MultiPointM = 28,
            /// <summary>Collection of surface patches</summary>
            MultiPatch = 31
        }

        /// <summary>
        /// Part type enumeration - everything but ShapeType.MultiPatch just uses PartType.Ring.
        /// </summary>
        public enum PartType
        {
            /// <summary>
            /// Linked strip of triangles, where every vertex (after the first two) completes a new triangle.
            /// A new triangle is always formed by connecting the new vertex with its two immediate predecessors.
            /// </summary>
            TriangleStrip = 0,
            /// <summary>
            /// A linked fan of triangles, where every vertex (after the first two) completes a new triangle.
            /// A new triangle is always formed by connecting the new vertex with its immediate predecessor 
            /// and the first vertex of the part.
            /// </summary>
            TriangleFan = 1,
            /// <summary>The outer ring of a polygon</summary>
            OuterRing = 2,
            /// <summary>The first ring of a polygon</summary>
            InnerRing = 3,
            /// <summary>The outer ring of a polygon of an unspecified type</summary>
            FirstRing = 4,
            /// <summary>A ring of a polygon of an unspecified type</summary>
            Ring = 5
        }

        /// <summary>
        /// SHPObject - represents on shape (without attributes) read from the .shp file.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public class SHPObject
        {
            ///<summary>Shape type as a ShapeType enum</summary>	
            public ShapeType shpType;
            ///<summary>Shape number (-1 is unknown/unassigned)</summary>	
            public int nShapeId;
            ///<summary>Number of parts (0 implies single part with no info)</summary>	
            public int nParts;
            ///<summary>Pointer to int array of part start offsets, of size nParts</summary>	
            public IntPtr paPartStart;
            ///<summary>Pointer to PartType array (PartType.Ring if not ShapeType.MultiPatch) of size nParts</summary>	
            public IntPtr paPartType;
            ///<summary>Number of vertices</summary>	
            public int nVertices;
            ///<summary>Pointer to double array containing X coordinates</summary>	
            public IntPtr padfX;
            ///<summary>Pointer to double array containing Y coordinates</summary>		
            public IntPtr padfY;
            ///<summary>Pointer to double array containing Z coordinates (all zero if not provided)</summary>	
            public IntPtr padfZ;
            ///<summary>Pointer to double array containing Measure coordinates(all zero if not provided)</summary>	
            public IntPtr padfM;
            ///<summary>Bounding rectangle's min X</summary>	
            public double dfXMin;
            ///<summary>Bounding rectangle's min Y</summary>	
            public double dfYMin;
            ///<summary>Bounding rectangle's min Z</summary>	
            public double dfZMin;
            ///<summary>Bounding rectangle's min M</summary>	
            public double dfMMin;
            ///<summary>Bounding rectangle's max X</summary>	
            public double dfXMax;
            ///<summary>Bounding rectangle's max Y</summary>	
            public double dfYMax;
            ///<summary>Bounding rectangle's max Z</summary>	
            public double dfZMax;
            ///<summary>Bounding rectangle's max M</summary>	
            public double dfMMax;
        }

        /// <summary>
        /// The SHPOpen() function should be used to establish access to the two files for 
        /// accessing vertices (.shp and .shx). Note that both files have to be in the indicated 
        /// directory, and must have the expected extensions in lower case. The returned SHPHandle 
        /// is passed to other access functions, and SHPClose() should be invoked to recover 
        /// resources, and flush changes to disk when complete.
        /// </summary>
        /// <param name="szShapeFile">The name of the layer to access.  This can be the name of either 
        /// the .shp or the .shx file or can just be the path plus the basename of the pair.</param>
        /// <param name="szAccess">The fopen() style access string. At this time only "rb" (read-only binary) 
        /// and "rb+" (read/write binary) should be used.</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPOpen", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPOpen64(string szShapeFile, string szAccess);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPOpen", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPOpen32(string szShapeFile, string szAccess);

        public static IntPtr SHPOpen(string szShapeFile, string szAccess)
        {
            if (IntPtr.Size == 4)
                return SHPOpen32(szShapeFile, szAccess);
            else
                return SHPOpen64(szShapeFile, szAccess);
        }

        /// <summary>
        /// The SHPCreate() function will create a new .shp and .shx file of the desired type.
        /// </summary>
        /// <param name="szShapeFile">The name of the layer to access. This can be the name of either 
        /// the .shp or the .shx file or can just be the path plus the basename of the pair.</param>
        /// <param name="shpType">The type of shapes to be stored in the newly created file. 
        /// It may be either ShapeType.Point, ShapeType.PolyLine, ShapeType.Polygon or ShapeType.MultiPoint.</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreate", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreate64(string szShapeFile, ShapeType shpType);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreate", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreate32(string szShapeFile, ShapeType shpType);

        public static IntPtr SHPCreate(string szShapeFile, ShapeType shpType)
        {
            if (IntPtr.Size == 4)
                return SHPCreate32(szShapeFile, shpType);
            else
                return SHPCreate64(szShapeFile, shpType);
        }

        /// <summary>
        /// The SHPGetInfo() function retrieves various information about shapefile as a whole. 
        /// The bounds are read from the file header, and may be inaccurate if the file was 
        /// improperly generated.
        /// </summary>
        /// <param name="hSHP">The handle previously returned by SHPOpen() or SHPCreate()</param>
        /// <param name="pnEntities">A pointer to an integer into which the number of 
        /// entities/structures should be placed. May be NULL.</param>
        /// <param name="pshpType">A pointer to an integer into which the ShapeType of this file 
        /// should be placed. Shapefiles may contain either ShapeType.Point, ShapeType.PolyLine, ShapeType.Polygon or 
        /// ShapeType.MultiPoint entities. This may be NULL.</param>
        /// <param name="adfMinBound">The X, Y, Z and M minimum values will be placed into this 
        /// four entry array. This may be NULL. </param>
        /// <param name="adfMaxBound">The X, Y, Z and M maximum values will be placed into this 
        /// four entry array. This may be NULL.</param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPGetInfo", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPGetInfo64(IntPtr hSHP, ref int pnEntities,
            ref ShapeType pshpType, double[] adfMinBound, double[] adfMaxBound);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPGetInfo", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPGetInfo32(IntPtr hSHP, ref int pnEntities,
            ref ShapeType pshpType, double[] adfMinBound, double[] adfMaxBound);

        public static void SHPGetInfo(IntPtr hSHP, ref int pnEntities,
            ref ShapeType pshpType, double[] adfMinBound, double[] adfMaxBound)
        {
            if (IntPtr.Size == 4)
                SHPGetInfo32(hSHP, ref pnEntities, ref pshpType, adfMinBound, adfMaxBound);
            else
                SHPGetInfo64(hSHP, ref pnEntities, ref pshpType, adfMinBound, adfMaxBound);
        }

        /// <summary>
        /// The SHPReadObject() call is used to read a single structure, or entity from the shapefile. 
        /// See the definition of the SHPObject structure for detailed information on fields of a SHPObject. 
        /// </summary>
        /// <param name="hSHP">The handle previously returned by SHPOpen() or SHPCreate().</param>
        /// <param name="iShape">The entity number of the shape to read. Entity numbers are between 0 
        /// and nEntities-1 (as returned by SHPGetInfo()).</param>
        /// <returns>IntPtr to a SHPObject.  Use System.Runtime.InteropServices.Marshal.PtrToStructure(...)
        /// to create an SHPObject copy in managed memory</returns>
        /// <remarks>
        /// IntPtr's returned from SHPReadObject() should be deallocated with SHPDestroyShape(). 
        /// SHPReadObject() will return IntPtr.Zero if an illegal iShape value is requested. 
        /// Note that the bounds placed into the SHPObject are those read from the file, and may not be correct. 
        /// For points the bounds are generated from the single point since bounds aren't normally provided 
        /// for point types. Generally the shapes returned will be of the type of the file as a whole. 
        /// However, any file may also contain type ShapeType.NullShape shapes which will have no geometry. 
        /// Generally speaking applications should skip rather than preserve them, as they usually 
        /// represented interactively deleted shapes.
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPReadObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPReadObject64(IntPtr hSHP, int iShape);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPReadObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPReadObject32(IntPtr hSHP, int iShape);

        public static IntPtr SHPReadObject(IntPtr hSHP, int iShape)
        {
            if (IntPtr.Size == 4)
                return SHPReadObject32(hSHP, iShape);
            else
                return SHPReadObject64(hSHP, iShape);
        }

        /// <summary>
        /// The SHPWriteObject() call is used to write a single structure, or entity to the shapefile. 
        /// See the definition of the SHPObject structure for detailed information on fields of a SHPObject.
        /// </summary>
        /// <param name="hSHP">The handle previously returned by SHPOpen("r+") or SHPCreate().</param>
        /// <param name="iShape">The entity number of the shape to write. 
        /// A value of -1 should be used for new shapes. </param>
        /// <param name="psObject">The shape to write to the file. This should have been created with SHPCreateObject(), 
        /// or SHPCreateSimpleObject().</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPWriteObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPWriteObject64(IntPtr hSHP, int iShape, IntPtr psObject);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPWriteObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPWriteObject32(IntPtr hSHP, int iShape, IntPtr psObject);

        public static int SHPWriteObject(IntPtr hSHP, int iShape, IntPtr psObject)
        {
            if (IntPtr.Size == 4)
                return SHPWriteObject32(hSHP, iShape, psObject);
            else
                return SHPWriteObject64(hSHP, iShape, psObject);
        }

        /// <summary>
        /// This function should be used to deallocate the resources associated with a SHPObject 
        /// when it is no longer needed, including those created with SHPCreateSimpleObject(), 
        /// SHPCreateObject() and returned from SHPReadObject().
        /// </summary>
        /// <param name="psObject">IntPtr of the SHPObject to deallocate memory.</param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPDestroyObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPDestroyObject64(IntPtr psObject);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPDestroyObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPDestroyObject32(IntPtr psObject);

        public static void SHPDestroyObject(IntPtr psObject)
        {
            if (IntPtr.Size == 4)
                SHPDestroyObject32(psObject);
            else
                SHPDestroyObject64(psObject);
        }

        /// <summary>
        /// This function will recompute the extents of this shape, replacing the existing values 
        /// of the dfXMin, dfYMin, dfZMin, dfMMin, dfXMax, dfYMax, dfZMax, and dfMMax values based 
        /// on the current set of vertices for the shape. This function is automatically called by 
        /// SHPCreateObject() but if the vertices of an existing object are altered it should be 
        /// called again to fix up the extents.
        /// </summary>
        /// <param name="psObject">An existing shape object to be updated in place.</param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPComputeExtents", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPComputeExtents64(IntPtr psObject);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPComputeExtents", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPComputeExtents32(IntPtr psObject);

        public static void SHPComputeExtents(IntPtr psObject)
        {
            if (IntPtr.Size == 4)
                SHPComputeExtents32(psObject);
            else
                SHPComputeExtents64(psObject);
        }

        /// <summary>
        /// The SHPCreateObject() function allows for the creation of objects (shapes). 
        /// This is normally used so that the SHPObject can be passed to SHPWriteObject() 
        /// to write it to the file.
        /// </summary>
        /// <param name="shpType">The ShapeType of the object to be created, such as ShapeType.Point, or ShapeType.Polygon.</param>
        /// <param name="nShapeId">The shapeid to be recorded with this shape.</param>
        /// <param name="nParts">The number of parts for this object. If this is zero for PolyLine, 
        /// or Polygon type objects, a single zero valued part will be created internally.</param>
        /// <param name="panPartStart">The list of zero based start vertices for the rings 
        /// (parts) in this object. The first should always be zero. This may be NULL if nParts is 0.</param>
        /// <param name="paPartType">The type of each of the parts. This is only meaningful for MultiPatch files. 
        /// For all other cases this may be NULL, and will be assumed to be PartType.Ring.</param>
        /// <param name="nVertices">The number of vertices being passed in padfX, padfY, and padfZ. </param>
        /// <param name="adfX">An array of nVertices X coordinates of the vertices for this object.</param>
        /// <param name="adfY">An array of nVertices Y coordinates of the vertices for this object.</param>
        /// <param name="adfZ">An array of nVertices Z coordinates of the vertices for this object. 
        /// This may be NULL in which case they are all assumed to be zero.</param>
        /// <param name="adfM">An array of nVertices M (measure values) of the vertices for this object. 
        /// This may be NULL in which case they are all assumed to be zero.</param>
        /// <returns>IntPtr to a SHPObject</returns>
        /// <remarks>
        /// The SHPDestroyObject() function should be used to free 
        /// resources associated with an object allocated with SHPCreateObject(). This function 
        /// computes a bounding box for the SHPObject from the given vertices.
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateObject64(ShapeType shpType, int nShapeId,
            int nParts, int[] panPartStart, PartType[] paPartType,
            int nVertices, double[] adfX, double[] adfY,
            double[] adfZ, double[] adfM);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateObject32(ShapeType shpType, int nShapeId,
            int nParts, int[] panPartStart, PartType[] paPartType,
            int nVertices, double[] adfX, double[] adfY,
            double[] adfZ, double[] adfM);

        public static IntPtr SHPCreateObject(ShapeType shpType, int nShapeId,
            int nParts, int[] panPartStart, PartType[] paPartType,
            int nVertices, double[] adfX, double[] adfY,
            double[] adfZ, double[] adfM)
        {
            if (IntPtr.Size == 4)
                return SHPCreateObject32(shpType, nShapeId, nParts, panPartStart, paPartType, nVertices, adfX, adfY, adfZ, adfM);
            else
                return SHPCreateObject64(shpType, nShapeId, nParts, panPartStart, paPartType, nVertices, adfX, adfY, adfZ, adfM);
        }

        /// <summary>
        /// The SHPCreateSimpleObject() function allows for the convenient creation of simple objects. 
        /// This is normally used so that the SHPObject can be passed to SHPWriteObject() to write it 
        /// to the file. The simple object creation API assumes an M (measure) value of zero for each vertex. 
        /// For complex objects (such as polygons) it is assumed that there is only one part, and that it 
        /// is of the default type (PartType.Ring). Use the SHPCreateObject() function for more sophisticated 
        /// objects. 
        /// </summary>
        /// <param name="shpType">The ShapeType of the object to be created, such as ShapeType.Point, or ShapeType.Polygon.</param>
        /// <param name="nVertices">The number of vertices being passed in padfX, padfY, and padfZ.</param>
        /// <param name="adfX">An array of nVertices X coordinates of the vertices for this object.</param>
        /// <param name="adfY">An array of nVertices Y coordinates of the vertices for this object.</param>
        /// <param name="adfZ">An array of nVertices Z coordinates of the vertices for this object. 
        /// This may be NULL in which case they are all assumed to be zero.</param>
        /// <returns>IntPtr to a SHPObject</returns>
        /// <remarks>
        /// The SHPDestroyObject() function should be used to free resources associated with an 
        /// object allocated with SHPCreateSimpleObject().
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateSimpleObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateSimpleObject64(ShapeType shpType, int nVertices,
            double[] adfX, double[] adfY, double[] adfZ);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateSimpleObject", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateSimpleObject32(ShapeType shpType, int nVertices,
            double[] adfX, double[] adfY, double[] adfZ);

        public static IntPtr SHPCreateSimpleObject(ShapeType shpType, int nVertices,
            double[] adfX, double[] adfY, double[] adfZ)
        {
            if (IntPtr.Size == 4)
                return SHPCreateSimpleObject32(shpType, nVertices, adfX, adfY, adfZ);
            else
                return SHPCreateSimpleObject64(shpType, nVertices, adfX, adfY, adfZ);
        }

        /// <summary>
        /// The SHPClose() function will close the .shp and .shx files, and flush all outstanding header 
        /// information to the files. It will also recover resources associated with the handle. 
        /// After this call the hSHP handle cannot be used again.
        /// </summary>
        /// <param name="hSHP">The handle previously returned by SHPOpen() or SHPCreate().</param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPClose", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPClose64(IntPtr hSHP);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPClose", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPClose32(IntPtr hSHP);

        public static void SHPClose(IntPtr hSHP)
        {
            if (IntPtr.Size == 4)
                SHPClose32(hSHP);
            else
                SHPClose64(hSHP);
        }

        /// <summary>
        /// Translates a ShapeType.* constant into a named shape type (Point, PointZ, Polygon, etc.)
        /// </summary>
        /// <param name="shpType">ShapeType enum</param>
        /// <returns>string</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTypeName", CallingConvention = CallingConvention.Cdecl)]
        private static extern string SHPTypeName64(ShapeType shpType);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTypeName", CallingConvention = CallingConvention.Cdecl)]
        private static extern string SHPTypeName32(ShapeType shpType);

        public static string SHPTypeName(ShapeType shpType)
        {
            if (IntPtr.Size == 4)
               return SHPTypeName32(shpType);
            else
               return SHPTypeName64(shpType);
        }

        /// <summary>
        /// Translates a PartType enum into a named part type (Ring, Inner Ring, etc.)
        /// </summary>
        /// <param name="partType">PartType enum</param>
        /// <returns>string</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPPartTypeName", CallingConvention = CallingConvention.Cdecl)]
        private static extern string SHPPartTypeName64(PartType partType);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPPartTypeName", CallingConvention = CallingConvention.Cdecl)]
        private static extern string SHPPartTypeName32(PartType partType);

        public static string SHPTypeName(PartType partType)
        {
            if (IntPtr.Size == 4)
                return SHPPartTypeName32(partType);
            else
                return SHPPartTypeName64(partType);
        }

        /* -------------------------------------------------------------------- */
        /*      Shape quadtree indexing API.                                    */
        /* -------------------------------------------------------------------- */

        /// <summary>
        /// Creates a quadtree index
        /// </summary>
        /// <param name="hSHP"></param>
        /// <param name="nDimension"></param>
        /// <param name="nMaxDepth"></param>
        /// <param name="adfBoundsMin"></param>
        /// <param name="adfBoundsMax"></param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateTree", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateTree64(IntPtr hSHP, int nDimension, int nMaxDepth,
            double[] adfBoundsMin, double[] adfBoundsMax);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCreateTree", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPCreateTree32(IntPtr hSHP, int nDimension, int nMaxDepth,
            double[] adfBoundsMin, double[] adfBoundsMax);

        public static IntPtr SHPCreateTree(IntPtr hSHP, int nDimension, int nMaxDepth,
            double[] adfBoundsMin, double[] adfBoundsMax)
        {
            if (IntPtr.Size == 4)
                return SHPCreateTree32(hSHP, nDimension, nMaxDepth, adfBoundsMin, adfBoundsMax);
            else
                return SHPCreateTree64(hSHP, nDimension, nMaxDepth, adfBoundsMin, adfBoundsMax);
        }

        /// <summary>
        /// Releases resources associated with quadtree
        /// </summary>
        /// <param name="hTree"></param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPDestroyTree", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPDestroyTree64(IntPtr hTree);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPDestroyTree", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPDestroyTree32(IntPtr hTree);

        public static void SHPDestroyTree(IntPtr hTree)
        {
            if (IntPtr.Size == 4)
                SHPDestroyTree32(hTree);
            else
                SHPDestroyTree64(hTree);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hTree"></param>
        /// <param name="psObject"></param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeAddShapeId", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPTreeAddShapeId64(IntPtr hTree, IntPtr psObject);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeAddShapeId", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPTreeAddShapeId32(IntPtr hTree, IntPtr psObject);

        public static int SHPTreeAddShapeId(IntPtr hTree, IntPtr psObject)
        {
            if (IntPtr.Size == 4)
                return SHPTreeAddShapeId32(hTree, psObject);
            else
                return SHPTreeAddShapeId64(hTree, psObject);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hTree"></param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeTrimExtraNodes", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPTreeTrimExtraNodes64(IntPtr hTree);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeTrimExtraNodes", CallingConvention = CallingConvention.Cdecl)]
        private static extern void SHPTreeTrimExtraNodes32(IntPtr hTree);

        public static void SHPTreeTrimExtraNodes(IntPtr hTree)
        {
            if (IntPtr.Size == 4)
                SHPTreeTrimExtraNodes32(hTree);
            else
                SHPTreeTrimExtraNodes64(hTree);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hTree"></param>
        /// <param name="adfBoundsMin"></param>
        /// <param name="adfBoundsMax"></param>
        /// <param name="pnShapeCount"></param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeFindLikelyShapes", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPTreeFindLikelyShapes64(IntPtr hTree,
            double[] adfBoundsMin, double[] adfBoundsMax, ref int pnShapeCount);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPTreeFindLikelyShapes", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr SHPTreeFindLikelyShapes32(IntPtr hTree,
            double[] adfBoundsMin, double[] adfBoundsMax, ref int pnShapeCount);

        public static IntPtr SHPTreeFindLikelyShapes(IntPtr hTree,
            double[] adfBoundsMin, double[] adfBoundsMax, ref int pnShapeCount)
        {
            if (IntPtr.Size == 4)
                return SHPTreeFindLikelyShapes32(hTree, adfBoundsMin, adfBoundsMax, ref pnShapeCount);
            else
                return SHPTreeFindLikelyShapes64(hTree, adfBoundsMin, adfBoundsMax, ref pnShapeCount);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adfBox1Min"></param>
        /// <param name="adfBox1Max"></param>
        /// <param name="adfBox2Min"></param>
        /// <param name="adfBox2Max"></param>
        /// <param name="nDimension"></param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCheckBoundsOverlap", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPCheckBoundsOverlap64(double[] adfBox1Min, double[] adfBox1Max,
            double[] adfBox2Min, double[] adfBox2Max, int nDimension);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "SHPCheckBoundsOverlap", CallingConvention = CallingConvention.Cdecl)]
        private static extern int SHPCheckBoundsOverlap32(double[] adfBox1Min, double[] adfBox1Max,
            double[] adfBox2Min, double[] adfBox2Max, int nDimension);

        public static int SHPCheckBoundsOverlap(double[] adfBox1Min, double[] adfBox1Max,
            double[] adfBox2Min, double[] adfBox2Max, int nDimension)
        {
            if (IntPtr.Size == 4)
                return SHPCheckBoundsOverlap32(adfBox1Min, adfBox1Max, adfBox2Min, adfBox2Max, nDimension);
            else
                return SHPCheckBoundsOverlap64(adfBox1Min, adfBox1Max, adfBox2Min, adfBox2Max, nDimension);
        }

        /// <summary>
        /// xBase字段
        /// </summary>
        public class DBFField
        {
            /// <summary>
            /// 字段名
            /// </summary>
            public string Name;
            /// <summary>
            /// 字段类型
            /// </summary>
            public DBFFieldType Type;
            /// <summary>
            /// 宽度
            /// </summary>
            public int Width;
            /// <summary>
            /// 小数位数
            /// </summary>
            public int Decimals;
            /// <summary>
            /// 字段描述
            /// </summary>
            public string Description;

            /// <summary>
            /// 新建字段
            /// </summary>
            /// <param name="name">字段名</param>
            /// <param name="type">字段类型</param>
            /// <param name="width">宽度</param>
            /// <param name="decimals">小数位数</param>
            /// <param name="description">字段描述</param>
            public DBFField(string name, DBFFieldType type, int width, int decimals, string description="")
            {
                Name = name;
                Type = type;
                Width = width;
                Decimals = decimals;
                Description = description;
            }
        }

        /// <summary>
        /// xBase field type enumeration
        /// </summary>
        public enum DBFFieldType
        {
            ///<summary>String data type</summary> 
            FTString,
            ///<summary>Integer data type</summary>
            FTInteger,
            ///<summary>Double data type</summary> 
            FTDouble,
            ///<summary>Logical data type</summary>
            FTLogical,
            ///<summary>Invalid data type</summary>
            FTInvalid,
            /// <summary>Date data type</summary>
            FTDate
        };

        /// <summary>
        /// The DBFOpen() function should be used to establish access to an existing xBase format table file. 
        /// The returned DBFHandle is passed to other access functions, and DBFClose() should be invoked 
        /// to recover resources, and flush changes to disk when complete. The DBFCreate() function should 
        /// called to create new xBase files. As a convenience, DBFOpen() can be called with the name of a 
        /// .shp or .shx file, and it will figure out the name of the related .dbf file.
        /// </summary>
        /// <param name="szDBFFile">The name of the xBase (.dbf) file to access.</param>
        /// <param name="szAccess">The fopen() style access string. At this time only "rb" (read-only binary) 
        /// and "rb+" (read/write binary) should be used.</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFOpen", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFOpen64(string szDBFFile, string szAccess);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFOpen", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFOpen32(string szDBFFile, string szAccess);

        /// <summary>
        /// 打开DBF文件
        /// </summary>
        /// <param name="szDBFFile">DBF文件名</param>
        /// <param name="szAccess">只读/读写（"rb"/"rb+"）</param>
        /// <returns>IntPtr</returns>
        public static IntPtr DBFOpen(string szDBFFile, string szAccess)
        {
            if (IntPtr.Size == 4)
                return DBFOpen32(szDBFFile, szAccess);
            else
                return DBFOpen64(szDBFFile, szAccess);
        }

        /// <summary>
        /// The DBFCreate() function creates a new xBase format file with the given name, 
        /// and returns an access handle that can be used with other DBF functions. 
        /// The newly created file will have no fields, and no records. 
        /// Fields should be added with DBFAddField() before any records add written. 
        /// </summary>
        /// <param name="szDBFFile">The name of the xBase (.dbf) file to create.</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFCreate", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFCreate64(string szDBFFile);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFCreate", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFCreate32(string szDBFFile);

        public static IntPtr DBFCreate(string szDBFFile)
        {
            if (IntPtr.Size == 4)
                return DBFCreate32(szDBFFile);
            else
                return DBFCreate64(szDBFFile);
        }

        /// <summary>
        /// The DBFGetFieldCount() function returns the number of fields currently defined 
        /// for the indicated xBase file. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldCount", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetFieldCount64(IntPtr hDBF);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldCount", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetFieldCount32(IntPtr hDBF);

        public static int DBFGetFieldCount(IntPtr hDBF)
        {
            if (IntPtr.Size == 4)
                return DBFGetFieldCount32(hDBF);
            else
                return DBFGetFieldCount64(hDBF);
        }

        /// <summary>
        /// The DBFGetRecordCount() function returns the number of records that exist on the xBase 
        /// file currently. Note that with shape files one xBase record exists for each shape in the 
        /// .shp/.shx files.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetRecordCount", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetRecordCount64(IntPtr hDBF);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetRecordCount", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetRecordCount32(IntPtr hDBF);

        public static int DBFGetRecordCount(IntPtr hDBF)
        {
            if (IntPtr.Size == 4)
                return DBFGetRecordCount32(hDBF);
            else
                return DBFGetRecordCount64(hDBF);
        }

        /// <summary>
        /// The DBFAddField() function is used to add new fields to an existing xBase file opened with DBFOpen(), 
        /// or created with DBFCreate(). Note that fields can only be added to xBase files with no records, 
        /// though this is limitation of this API, not of the file format. Returns the field number of the 
        /// new field, or -1 if the addition of the field failed
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be updated, as returned by DBFOpen(), 
        /// or DBFCreate().</param>
        /// <param name="szFieldName">The name of the new field. At most 11 character will be used. 
        /// In order to use the xBase file in some packages it may be necessary to avoid some special 
        /// characters in the field names such as spaces, or arithmetic operators.</param>
        /// <param name="eType">One of FTString, FTInteger, FTLogical, FTDate, or FTDouble in order to establish the 
        /// type of the new field. Note that some valid xBase field types cannot be created such as date fields.</param>
        /// <param name="nWidth">The width of the field to be created. For FTString fields this establishes 
        /// the maximum length of string that can be stored. For FTInteger this establishes the number of 
        /// digits of the largest number that can be represented. For FTDouble fields this in combination 
        /// with the nDecimals value establish the size, and precision of the created field.</param>
        /// <param name="nDecimals">The number of decimal places to reserve for FTDouble fields. 
        /// For all other field types this should be zero. For instance with nWidth=7, and nDecimals=3 
        /// numbers would be formatted similarly to `123.456'.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFAddField", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFAddField64(IntPtr hDBF, string szFieldName,
            DBFFieldType eType, int nWidth, int nDecimals);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFAddField", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFAddField32(IntPtr hDBF, string szFieldName,
            DBFFieldType eType, int nWidth, int nDecimals);

        public static int DBFAddField(IntPtr hDBF, string szFieldName,
            DBFFieldType eType, int nWidth, int nDecimals)
        {
            if (IntPtr.Size == 4)
                return DBFAddField32(hDBF, szFieldName, eType, nWidth, nDecimals);
            else
                return DBFAddField64(hDBF, szFieldName, eType, nWidth, nDecimals);
        }

        /// <summary>
        /// The DBFGetFieldInfo() returns the type of the requested field, which is one of the DBFFieldType 
        /// enumerated values. As well, the field name, and field width information can optionally be returned. 
        /// The field type returned does not correspond one to one with the xBase field types. 
        /// For instance the xBase field type for Date will just be returned as being FTInteger. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by DBFOpen(), 
        /// or DBFCreate().</param>
        /// <param name="iField">The field to be queried. This should be a number between 0 and n-1, 
        /// where n is the number fields on the file, as returned by DBFGetFieldCount().</param>
        /// <param name="szFieldName">If this pointer is not NULL the name of the requested field 
        /// will be written to this location. The pszFieldName buffer should be at least 12 character 
        /// is size in order to hold the longest possible field name of 11 characters plus a terminating 
        /// zero character.</param>
        /// <param name="pnWidth">If this pointer is not NULL, the width of the requested field will be 
        /// returned in the int pointed to by pnWidth. This is the width in characters. </param>
        /// <param name="pnDecimals">If this pointer is not NULL, the number of decimal places precision 
        /// defined for the field will be returned. This is zero for integer fields, or non-numeric fields.</param>
        /// <returns>DBFFieldType</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldInfo", CallingConvention = CallingConvention.Cdecl)]
        private static extern DBFFieldType DBFGetFieldInfo64(IntPtr hDBF, int iField,
            StringBuilder szFieldName, ref int pnWidth, ref int pnDecimals);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldInfo", CallingConvention = CallingConvention.Cdecl)]
        private static extern DBFFieldType DBFGetFieldInfo32(IntPtr hDBF, int iField,
            StringBuilder szFieldName, ref int pnWidth, ref int pnDecimals);

        public static DBFFieldType DBFGetFieldInfo(IntPtr hDBF, int iField,
            StringBuilder szFieldName, ref int pnWidth, ref int pnDecimals)
        {
            if (IntPtr.Size == 4)
                return DBFGetFieldInfo32(hDBF, iField, szFieldName, ref pnWidth, ref pnDecimals);
            else
                return DBFGetFieldInfo64(hDBF, iField, szFieldName, ref pnWidth, ref pnDecimals);
        }

        /// <summary>
        /// Returns the index of the field matching this name, or -1 on failure. 
        /// The comparison is case insensitive. However, lengths must match exactly.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned 
        /// by DBFOpen(), or DBFCreate().</param>
        /// <param name="szFieldName">Name of the field to search for.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldIndex", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetFieldIndex64(IntPtr hDBF, string szFieldName);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetFieldIndex", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFGetFieldIndex32(IntPtr hDBF, string szFieldName);

        public static int DBFGetFieldIndex(IntPtr hDBF, string szFieldName)
        {
            if (IntPtr.Size == 4)
                return DBFGetFieldIndex32(hDBF, szFieldName);
            else
                return DBFGetFieldIndex64(hDBF, szFieldName);
        }

        /// <summary>
        /// The DBFReadIntegerAttribute() will read the value of one field and return it as an integer. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>int</returns>
        /// <remarks>
        /// This can be used even with FTString fields, though the returned value will be zero if not 
        /// interpretable as a number.
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadIntegerAttribute", CallingConvention = CallingConvention.Cdecl)]
        public static extern int DBFReadIntegerAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadIntegerAttribute", CallingConvention = CallingConvention.Cdecl)]
        public static extern int DBFReadIntegerAttribute32(IntPtr hDBF, int iShape, int iField);

        public static int DBFReadIntegerAttribute(IntPtr hDBF, int iShape, int iField)
        {
            if (IntPtr.Size == 4)
                return DBFReadIntegerAttribute32(hDBF, iShape, iField);
            else
                return DBFReadIntegerAttribute64(hDBF, iShape, iField);
        }

        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadDateAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFReadDateAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadDateAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFReadDateAttribute32(IntPtr hDBF, int iShape, int iField);
        /// <summary>
        /// The DBFReadDateAttribute() will read the value of one field and return it as a System.DateTime value.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>int</returns>
        /// <remarks>
        /// This can be used even with FTString fields, though the returned value will be zero if not 
        /// interpretable as a number.
        /// </remarks>
        public static DateTime DBFReadDateAttribute(IntPtr hDBF, int iShape, int iField)
        {
            string sDate;

            if (IntPtr.Size == 4)
                sDate = DBFReadDateAttribute32(hDBF, iShape, iField).ToString();
            else
                sDate = DBFReadDateAttribute64(hDBF, iShape, iField).ToString();

            try
            {
                DateTime d = new DateTime(int.Parse(sDate.Substring(0, 4)), int.Parse(sDate.Substring(4, 2)),
                    int.Parse(sDate.Substring(6, 2)));
                return d;
            }
            catch
            {
                return new DateTime(0);
            }
        }

        /// <summary>
        /// The DBFReadDoubleAttribute() will read the value of one field and return it as a double. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>double</returns>
        /// <remarks>
        /// This can be used even with FTString fields, though the returned value will be zero if not 
        /// interpretable as a number.
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadDoubleAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern double DBFReadDoubleAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadDoubleAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern double DBFReadDoubleAttribute32(IntPtr hDBF, int iShape, int iField);

        public static double DBFReadDoubleAttribute(IntPtr hDBF, int iShape, int iField)
        {
            if (IntPtr.Size == 4)
                return DBFReadDoubleAttribute32(hDBF, iShape, iField);
            else
                return DBFReadDoubleAttribute64(hDBF, iShape, iField);
        }

        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadStringAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFReadStringAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadStringAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFReadStringAttribute32(IntPtr hDBF, int iShape, int iField);
        /// <summary>
        /// The DBFReadStringAttribute() will read the value of one field and return it as a string. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>string</returns>
        /// <remarks>
        /// This function may be used on any field type (including FTInteger and FTDouble) and will 
        /// return the string representation stored in the .dbf file. The returned pointer is to an 
        /// internal buffer which is only valid untill the next DBF function call. It's contents may 
        /// be copied with normal string functions such as strcpy(), or strdup(). If the 
        /// TRIM_DBF_WHITESPACE macro is defined in shapefil.h (it is by default) then all leading and 
        /// trailing space (ASCII 32) characters will be stripped before the string is returned.
        /// </remarks>
        public static string DBFReadStringAttribute(IntPtr hDBF, int iShape, int iField)
        {
            IntPtr stringAnsi;

            if (IntPtr.Size == 4)
                stringAnsi = DBFReadStringAttribute32(hDBF, iShape, iField);
            else
                stringAnsi = DBFReadStringAttribute64(hDBF, iShape, iField);

            if (stringAnsi == IntPtr.Zero) { return string.Empty; }
            var stringUni = Marshal.PtrToStringUni(stringAnsi, 255);
            if (string.IsNullOrEmpty(stringUni)) { return string.Empty; }

            var bytes = Encoding.Unicode.GetBytes(stringUni);
            var attribute = Encoding.UTF8.GetString(bytes);
            var index = attribute.IndexOf('\0');
            return index > 0 ? attribute.Substring(0, index).Trim() : string.Empty;
        }

        public static string DBFReadAnsiStringAttribute(IntPtr hDBF, int iShape, int iField)
        {
            IntPtr stringAnsi;

            if (IntPtr.Size == 4)
                stringAnsi = DBFReadStringAttribute32(hDBF, iShape, iField);
            else
                stringAnsi = DBFReadStringAttribute64(hDBF, iShape, iField);

            if (stringAnsi == IntPtr.Zero) { return string.Empty; }
            var attributeAnsi = Marshal.PtrToStringAnsi(stringAnsi);
            return attributeAnsi;
        }

        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadLogicalAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern string DBFReadLogicalAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadLogicalAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern string DBFReadLogicalAttribute32(IntPtr hDBF, int iShape, int iField);
        /// <summary>
        /// The DBFReadLogicalAttribute() will read the value of one field and return it as a boolean. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>bool</returns>
        /// <remarks>
        /// This can be used with FTString fields, in which case it returns TRUE if the string="T";
        /// otherwise it returns FALSE.
        /// </remarks>
        public static bool DBFReadLogicalAttribute(IntPtr hDBF, int iShape, int iField)
        {
            if (IntPtr.Size == 4)
                return (DBFReadLogicalAttribute32(hDBF, iShape, iField) == "T");
            else
                return (DBFReadLogicalAttribute64(hDBF, iShape, iField) == "T");
        }

        /// <summary>
        /// This function will return TRUE if the indicated field is NULL valued otherwise FALSE. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) from which the field value should be read.</param>
        /// <param name="iField">The field within the selected record that should be read.</param>
        /// <returns>int</returns>
        /// <remarks>
        /// Note that NULL fields are represented in the .dbf file as having all spaces in the field. 
        /// Reading NULL fields will result in a value of 0.0 or an empty string with the other 
        /// DBFRead*Attribute() functions.
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFIsAttributeNULL", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFIsAttributeNULL64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFIsAttributeNULL", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFIsAttributeNULL32(IntPtr hDBF, int iShape, int iField);

        public static int DBFIsAttributeNULL(IntPtr hDBF, int iShape, int iField)
        {
            if (IntPtr.Size == 4)
                return DBFIsAttributeNULL32(hDBF, iShape, iField);
            else
                return DBFIsAttributeNULL64(hDBF, iShape, iField);
        }

        /// <summary>
        /// The DBFWriteIntegerAttribute() function is used to write a value to a numeric field 
        /// (FTInteger, or FTDouble). If the write succeeds the value TRUE will be returned, 
        /// otherwise FALSE will be returned. If the value is too large to fit in the field, 
        /// it will be truncated and FALSE returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="nFieldValue">The integer value that should be written.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteIntegerAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteIntegerAttribute64(IntPtr hDBF, int iShape,
            int iField, int nFieldValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteIntegerAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteIntegerAttribute32(IntPtr hDBF, int iShape,
            int iField, int nFieldValue);

        public static int DBFWriteIntegerAttribute(IntPtr hDBF, int iShape,
            int iField, int nFieldValue)
        {
            if (IntPtr.Size == 4)
                return DBFWriteIntegerAttribute32(hDBF, iShape, iField, nFieldValue);
            else
                return DBFWriteIntegerAttribute64(hDBF, iShape, iField, nFieldValue);
        }

        /// <summary>
        /// Write an attribute record to the file, but without any reformatting based on type.  
        /// The provided buffer is written as is to the field position in the record. 
        /// If the write succeeds the value TRUE will be returned, otherwise FALSE will be returned. 
        /// If the value is too large to fit in the field, it will be truncated and FALSE returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="pValue">pointer to the value</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteAttributeDirectly", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteAttributeDirectly64(IntPtr hDBF, int iShape,
            int iField, IntPtr pValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteAttributeDirectly", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteAttributeDirectly32(IntPtr hDBF, int iShape,
            int iField, IntPtr pValue);

        public static int DBFWriteAttributeDirectly(IntPtr hDBF, int iShape,
            int iField, IntPtr pValue)
        {
            if (IntPtr.Size == 4)
                return DBFWriteAttributeDirectly32(hDBF, iShape, iField, pValue);
            else
                return DBFWriteAttributeDirectly64(hDBF, iShape, iField, pValue);
        }


        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteDateAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteDateAttribute64(IntPtr hDBF, int iShape,
            int iField, int nFieldValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteDateAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteDateAttribute32(IntPtr hDBF, int iShape,
            int iField, int nFieldValue);
        /// <summary>
        /// The DBFWriteDateAttribute() function is used to write a value to a date field 
        /// (FTDate). If the write succeeds the value TRUE will be returned, 
        /// otherwise FALSE will be returned. If the value is too large to fit in the field, 
        /// it will be truncated and FALSE returned. 
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="nFieldValue">The System.DateTime date value that should be written.</param>
        /// <returns>int</returns>
        public static int DBFWriteDateAttribute(IntPtr hDBF, int iShape,
            int iField, DateTime nFieldValue)
        {
            if (IntPtr.Size == 4)
                return DBFWriteDateAttribute32(hDBF, iShape, iField, int.Parse(nFieldValue.ToString("yyyyMMdd")));
            else
                return DBFWriteDateAttribute64(hDBF, iShape, iField, int.Parse(nFieldValue.ToString("yyyyMMdd")));
        }

        /// <summary>
        /// The DBFWriteDoubleAttribute() function is used to write a value to a numeric field 
        /// (FTInteger, or FTDouble). If the write succeeds the value TRUE will be returned, 
        /// otherwise FALSE will be returned. If the value is too large to fit in the field, 
        /// it will be truncated and FALSE returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="dFieldValue">The floating point value that should be written.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteDoubleAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteDoubleAttribute64(IntPtr hDBF, int iShape,
            int iField, double dFieldValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteDoubleAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteDoubleAttribute32(IntPtr hDBF, int iShape,
            int iField, double dFieldValue);

        public static int DBFWriteDoubleAttribute(IntPtr hDBF, int iShape,
            int iField, double dFieldValue)
        {
            if (IntPtr.Size == 4)
                return DBFWriteDoubleAttribute32(hDBF, iShape, iField, dFieldValue);
            else
                return DBFWriteDoubleAttribute64(hDBF, iShape, iField, dFieldValue);
        }

        /// <summary>
        /// The DBFWriteStringAttribute() function is used to write a value to a string field (FString). 
        /// If the write succeeds the value TRUE willbe returned, otherwise FALSE will be returned. 
        /// If the value is too large to fit in the field, it will be truncated and FALSE returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="szFieldValue">The string to be written to the field.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteStringAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteStringAttribute64(IntPtr hDBF, int iShape,
            int iField, string szFieldValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteStringAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteStringAttribute32(IntPtr hDBF, int iShape,
            int iField, string szFieldValue);

        public static int DBFWriteStringAttribute(IntPtr hDBF, int iShape,
            int iField, string szFieldValue)
        {
            if (IntPtr.Size == 4)
                return DBFWriteStringAttribute32(hDBF, iShape, iField, szFieldValue);
            else
                return DBFWriteStringAttribute64(hDBF, iShape, iField, szFieldValue);
        }

        /// <summary>
        /// The DBFWriteNULLAttribute() function is used to clear the indicated field to a NULL value. 
        /// In the .dbf file this is represented by setting the entire field to spaces. If the write 
        /// succeeds the value TRUE will be returned, otherwise FALSE will be returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteNULLAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteNULLAttribute64(IntPtr hDBF, int iShape, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteNULLAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteNULLAttribute32(IntPtr hDBF, int iShape, int iField);

        public static int DBFWriteNULLAttribute(IntPtr hDBF, int iShape, int iField)
        {
            if (IntPtr.Size == 4)
                return DBFWriteNULLAttribute32(hDBF, iShape, iField);
            else
                return DBFWriteNULLAttribute64(hDBF, iShape, iField);
        }

        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteLogicalAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteLogicalAttribute64(IntPtr hDBF, int iShape,
            int iField, char lFieldValue);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteLogicalAttribute", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteLogicalAttribute32(IntPtr hDBF, int iShape,
            int iField, char lFieldValue);
        /// <summary>
        /// The DBFWriteLogicalAttribute() function is used to write a boolean value to a logical field 
        /// (FTLogical). If the write succeeds the value TRUE will be returned, 
        /// otherwise FALSE will be returned.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be written, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="iShape">The record number (shape number) to which the field value should be written.</param>
        /// <param name="iField">The field within the selected record that should be written.</param>
        /// <param name="bFieldValue">The boolean value to be written to the field.</param>
        /// <returns>int</returns>
        public static int DBFWriteLogicalAttribute(IntPtr hDBF, int iShape, int iField, bool bFieldValue)
        {
            if (IntPtr.Size == 4)
            {
                if (bFieldValue)
                    return DBFWriteLogicalAttribute32(hDBF, iShape, iField, 'T');
                else
                    return DBFWriteLogicalAttribute32(hDBF, iShape, iField, 'F');
            }
            else
            {
                if (bFieldValue)
                    return DBFWriteLogicalAttribute64(hDBF, iShape, iField, 'T');
                else
                    return DBFWriteLogicalAttribute64(hDBF, iShape, iField, 'F');
            }
        }

        /// <summary>
        /// Reads the attribute fields of a record.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="hEntity">The entity (record) number to be read</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadTuple", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFReadTuple64(IntPtr hDBF, int hEntity);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFReadTuple", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFReadTuple32(IntPtr hDBF, int hEntity);

        public static IntPtr DBFReadTuple(IntPtr hDBF, int hEntity)
        {
            if (IntPtr.Size == 4)
                return DBFReadTuple32(hDBF, hEntity);
            else
                return DBFReadTuple64(hDBF, hEntity);
        }

        /// <summary>
        /// Writes an attribute record to the file.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="hEntity">The zero-based entity (record) number to be written.  If hEntity equals 
        /// the number of records a new record is appended.</param>
        /// <param name="pRawTuple">Pointer to the tuple to be written</param>
        /// <returns>int</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteTuple", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteTuple64(IntPtr hDBF, int hEntity, IntPtr pRawTuple);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFWriteTuple", CallingConvention = CallingConvention.Cdecl)]
        private static extern int DBFWriteTuple32(IntPtr hDBF, int hEntity, IntPtr pRawTuple);

        public static int DBFWriteTuple(IntPtr hDBF, int hEntity, IntPtr pRawTuple)
        {
            if (IntPtr.Size == 4)
                return DBFWriteTuple32(hDBF, hEntity, pRawTuple);
            else
                return DBFWriteTuple64(hDBF, hEntity, pRawTuple);
        }

        /// <summary>
        /// Copies the data structure of an xBase file to another xBase file.  
        /// Data are not copied.  Use Read/WriteTuple functions to selectively copy data.
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be queried, as returned by 
        /// DBFOpen(), or DBFCreate().</param>
        /// <param name="szFilename">The name of the xBase (.dbf) file to create.</param>
        /// <returns>IntPtr</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFCloneEmpty", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFCloneEmpty64(IntPtr hDBF, string szFilename);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFCloneEmpty", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr DBFCloneEmpty32(IntPtr hDBF, string szFilename);

        public static IntPtr DBFCloneEmpty(IntPtr hDBF, string szFilename)
        {
            if (IntPtr.Size == 4)
                return DBFCloneEmpty32(hDBF, szFilename);
            else
                return DBFCloneEmpty64(hDBF, szFilename);
        }

        /// <summary>
        /// The DBFClose() function will close the indicated xBase file (opened with DBFOpen(), 
        /// or DBFCreate()), flushing out all information to the file on disk, and recovering 
        /// any resources associated with having the file open. The file handle (hDBF) should not 
        /// be used again with the DBF API after calling DBFClose().
        /// </summary>
        /// <param name="hDBF">The access handle for the file to be closed.</param>
        /// <returns>void</returns>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFClose", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DBFClose64(IntPtr hDBF);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFClose", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DBFClose32(IntPtr hDBF);

        public static void DBFClose(IntPtr hDBF)
        {
            if (IntPtr.Size == 4)
                DBFClose32(hDBF);
            else
                DBFClose64(hDBF);
        }


        /// <summary>
        /// This function returns the DBF type code of the indicated field.
        /// </summary>
        /// <param name="hDBF">The access handle for the file.</param>
        /// <param name="iField">The field index to query.</param>
        /// <returns>sbyte</returns>
        /// <remarks>
        /// Return value will be one of:
        /// <list type="bullet">
        /// <item><term>C</term><description>String</description></item>
        /// <item><term>D</term><description>Date</description></item>
        /// <item><term>F</term><description>Float</description></item>
        /// <item><term>N</term><description>Numeric, with or without decimal</description></item>
        /// <item><term>L</term><description>Logical</description></item>
        /// <item><term>M</term><description>Memo: 10 digits .DBT block ptr</description></item>
        /// <item><term> </term><description>field out of range</description></item>
        /// </list>
        /// </remarks>
        [DllImport("shapelib_x64.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetNativeFieldType", CallingConvention = CallingConvention.Cdecl)]
        private static extern sbyte DBFGetNativeFieldType64(IntPtr hDBF, int iField);
        [DllImport("shapelib_x86.dll", CharSet = CharSet.Ansi, EntryPoint = "DBFGetNativeFieldType", CallingConvention = CallingConvention.Cdecl)]
        private static extern sbyte DBFGetNativeFieldType32(IntPtr hDBF, int iField);

        public static sbyte DBFGetNativeFieldType(IntPtr hDBF, int iField)
        {
            if (IntPtr.Size == 4)
                return DBFGetNativeFieldType32(hDBF, iField);
            else
                return DBFGetNativeFieldType64(hDBF, iField);
        }

        /// <summary>
        /// private constructor:  no instantiation needed or permitted
        /// </summary>
        private Shapelib() { }

        /// <summary>
        /// 计算Polygon面积
        /// </summary>
        /// <param name="padfX"></param>
        /// <param name="padfY"></param>
        /// <returns></returns>
        public static double SHPCalculateArea(IntPtr hSHP, int iShape)
        {
            // SHPObject对象
            Shapelib.SHPObject shpObject = new Shapelib.SHPObject();
            // 读取SHPObject对象指针并将其转换为SHPObject对象
            Marshal.PtrToStructure(Shapelib.SHPReadObject(hSHP, iShape), shpObject);

            // 顶点数
            int nVertices = shpObject.nVertices;

            // 顶点的X坐标数组
            var padfX = new double[nVertices];
            // 将顶点的X坐标数组指针转换为数组
            Marshal.Copy(shpObject.padfX, padfX, 0, nVertices);
            // 顶点的Y坐标数组
            var padfY = new double[nVertices];
            // 将顶点的Y坐标数组指针转换为数组
            Marshal.Copy(shpObject.padfY, padfY, 0, nVertices);

            if (padfX.Length<=0 || padfY.Length<=0 || padfX.Length!= padfY.Length) { return 0; }

            var points = new List<Coordinate>();
            for (int i = 0; i < padfX.Length; i++) { points.Add(new Coordinate(padfY[i], padfX[i])); }
            double entityArea = Math.Abs(points.Take(points.Count - 1)
                        .Select((p, i) => (points[i + 1].X - p.X) * (points[i + 1].Y + p.Y))
                        .Sum() / 2);

            return entityArea;
        }

        /// <summary>
        /// 生成报备坐标文件
        /// </summary>
        /// <param name="coordinateFileName">报备坐标文件名</param>
        /// <param name="coordinateProperty">坐标属性描述</param>
        /// <param name="shpFileName">shp文件名</param>
        /// <param name="identifierFieldName">地块编号属性字段名</param>
        public static void SHPCreateCoordinateFile(string coordinateFileName, CoordinateProperty coordinateProperty, string shpFileName, string identifierFieldName)
        {
            var pathName = Path.GetDirectoryName(shpFileName);
            var fileName = Path.GetFileNameWithoutExtension(shpFileName);

            var shp = string.Format(@"{0}\{1}.shp", pathName, fileName);
            var hSHP = SHPOpen(shp, "rb");
            var dbf = string.Format(@"{0}\{1}.dbf", pathName, fileName);
            var hDBF = DBFOpen(dbf, "rb");

            // 实体数
            int pnEntities = 0;
            // 形状类型
            ShapeType pshpType = ShapeType.NullShape;
            // 界限坐标数组
            double[] adfMinBound = new double[4], adfMaxBound = new double[4];

            // 获取实体数、形状类型、界限坐标等信息
            Shapelib.SHPGetInfo(hSHP, ref pnEntities, ref pshpType, adfMinBound, adfMaxBound);

            int JIndex = 1;

            FileStream fs = new FileStream(coordinateFileName, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter m_streamWriter = new StreamWriter(fs);
            m_streamWriter.Flush();
            // 使用StreamWriter向文件中写入内容
            m_streamWriter.BaseStream.Seek(0, SeekOrigin.Begin);
            // 写入文件头
            m_streamWriter.WriteLine("[属性描述]");
            m_streamWriter.WriteLine(string.Format("坐标系={0}", coordinateProperty.CoordinateSystem));
            m_streamWriter.WriteLine(string.Format("几度分带={0}", coordinateProperty.ZoneType));
            m_streamWriter.WriteLine(string.Format("投影类型={0}", coordinateProperty.ProjectionType));
            m_streamWriter.WriteLine(string.Format("计量单位={0}", coordinateProperty.Unit));
            m_streamWriter.WriteLine(string.Format("带号={0}", coordinateProperty.Zone));
            m_streamWriter.WriteLine(string.Format("精度={0}", coordinateProperty.Decimals));
            m_streamWriter.WriteLine(string.Format("转换参数={0}", coordinateProperty.Parameters));
            m_streamWriter.WriteLine("[地块坐标]");

            // pnEntities
            for (int iShape = 0; iShape < pnEntities; iShape++)
            {

                // SHPObject对象
                SHPObject shpObject = new SHPObject();

                // 读取SHPObject对象指针
                var shpObjectPtr = SHPReadObject(hSHP, iShape);

                // 忽略可能存在问题的实体
                if (shpObjectPtr == IntPtr.Zero) { continue; }
                // 指针转换为SHPObject对象
                Marshal.PtrToStructure(shpObjectPtr, shpObject);

                // 顶点数
                int nVertices = shpObject.nVertices;

                // 顶点的X坐标数组
                var padfX = new double[nVertices];
                // 将顶点的X坐标数组指针转换为数组
                Marshal.Copy(shpObject.padfX, padfX, 0, nVertices);
                // 顶点的Y坐标数组
                var padfY = new double[nVertices];
                // 将顶点的Y坐标数组指针转换为数组
                Marshal.Copy(shpObject.padfY, padfY, 0, nVertices);

                int iField = DBFGetFieldIndex(hDBF, identifierFieldName);
                string entityName = DBFReadStringAttribute(hDBF, iShape, iField);

                var points = new List<Coordinate>();
                for (int i = 0; i < padfX.Length; i++) { points.Add(new Coordinate(padfY[i], padfX[i])); }
                double entityArea = Math.Abs(points.Take(points.Count - 1)
                            .Select((p, i) => (points[i + 1].X - p.X) * (points[i + 1].Y + p.Y))
                            .Sum() / 2);

                // 界址点数,地块面积,地块编号,地块名称,记录图形属性(点、线、面),图幅号,地块用途,地类编码,@
                m_streamWriter.WriteLine(string.Format("{0},{1:0.0000},{2},{3},面,,,,@", nVertices - 1, Math.Round(entityArea / 10000d, 4), iShape + 1, entityName));
                // 写入坐标
                for (int i = 0; i < nVertices - 1; i++)
                {
                    m_streamWriter.WriteLine(string.Format("J{0},1,{1:0.000},{2:0.000}", JIndex, padfY[i], padfX[i]));
                    JIndex++;
                }
            }
            //关闭此文件
            m_streamWriter.Flush();
            m_streamWriter.Close();

            DBFClose(hDBF);
            SHPClose(hSHP);
        }

        /// <summary>
        /// 根据报备坐标文件生成shp
        /// </summary>
        /// <param name="coordinateFileName">报备坐标文件名</param>
        /// <param name="shpFileName">shp文件名</param>
        /// <param name="identifierFieldName">地块编号属性字段名</param>
        public static void SHPCreateFromCoordinateFile(string coordinateFileName, string shpFileName, string identifierFieldName)
        {
            var lines = File.ReadAllLines(coordinateFileName);

            var polygonLineIndexs = new List<int>();
            for (int i = 9; i < lines.Length; i++)
            {
                if (lines[i].EndsWith("@"))
                {
                    polygonLineIndexs.Add(i);
                }
            }

            var polygons = new List<PolygonX>();
            for (int i = 0; i < polygonLineIndexs.Count; i++)
            {
                var polygon = new PolygonX();
                polygon.Name = lines[polygonLineIndexs[i]].Split(',')[3];
                polygon.AreaInRecord = Convert.ToDouble(lines[polygonLineIndexs[i]].Split(',')[1]);

                var endLineNo = i == polygonLineIndexs.Count - 1 ? lines.Length : polygonLineIndexs[i + 1];

                for (int j = polygonLineIndexs[i] + 1; j < endLineNo; j++)
                {
                    var contents = lines[j].Split(',');
                    var coordinate = new Coordinate();
                    coordinate.X = Convert.ToDouble(contents[2]);
                    coordinate.Y = Convert.ToDouble(contents[3]);
                    polygon.Coordinates.Add(coordinate);
                }
                polygon.Coordinates.Add(polygon.Coordinates[0]);

                polygons.Add(polygon);
            }

            if (polygons.Count > 0)
            {
                var nVertices = polygons.Count;

                var hSHP = SHPCreate(shpFileName, ShapeType.Polygon);
                var dbfFileName = shpFileName.Replace(".shp", ".dbf");
                var hDBF = DBFCreate(dbfFileName);
                var iField = DBFAddField(hDBF, identifierFieldName, DBFFieldType.FTString, 30, 0);
                var iFieldArea = DBFAddField(hDBF, "记录面积", DBFFieldType.FTDouble, 20, 4);

                for (int i = 0; i < nVertices; i++)
                {
                    var count = polygons[i].Coordinates.Count;
                    var adfX = new double[count];
                    var adfY = new double[count];
                    var adfZ = new double[count];

                    for (int j = 0; j < count; j++)
                    {
                        adfX[j] = polygons[i].Coordinates[j].Y;
                        adfY[j] = polygons[i].Coordinates[j].X;
                        adfZ[j] = 0;
                    }

                    var psObject = SHPCreateSimpleObject(ShapeType.Polygon, count, adfX, adfY, adfZ);
                    var iPolygon = SHPWriteObject(hSHP, -1, psObject);
                    DBFWriteStringAttribute(hDBF, i, iField, polygons[i].Name);
                    DBFWriteDoubleAttribute(hDBF, i, iFieldArea, polygons[i].AreaInRecord);
                }
                DBFClose(hDBF);
                SHPClose(hSHP);
            }
        }
    }

    public class CoordinateProperty
    {
        /// <summary>
        /// 坐标系统
        /// </summary>
        public string CoordinateSystem;
        /// <summary>
        /// 3°或6°分带
        /// </summary>
        public short ZoneType;
        /// <summary>
        /// 投影类型
        /// </summary>
        public string ProjectionType;
        /// <summary>
        /// 计量单位
        /// </summary>
        public string Unit;
        /// <summary>
        /// 带号
        /// </summary>
        public short Zone;
        /// <summary>
        /// 精度
        /// </summary>
        public double Decimals;
        /// <summary>
        /// 转换参数
        /// </summary>
        public string Parameters = ",,,,,,";
    }

    public class Coordinate
    {
        public double X;
        public double Y;

        public Coordinate()
        {

        }

        public Coordinate(double x, double y)
        {
            X = x;
            Y = y;
        }
    }

    public class PolygonX
    {
        public string Name;
        public double AreaInRecord;
        public List<Coordinate> Coordinates;

        public PolygonX()
        {
            Coordinates = new List<Coordinate>();
        }
    }
}
