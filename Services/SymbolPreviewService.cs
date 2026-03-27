using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using AcadColors = Autodesk.AutoCAD.Colors;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using AcadGeom = Autodesk.AutoCAD.Geometry;
using DrawingImage = System.Drawing.Image;
using Plant3D.ProjectRuntimePalettes.Models;
using Plant3D.ProjectRuntimePalettes.Utilities;

namespace Plant3D.ProjectRuntimePalettes.Services;

public sealed class SymbolPreviewService : IDisposable
{
    public const int PreviewPixelSize = 104;
    private const int PreviewSize = PreviewPixelSize;

    private readonly ProjectStyleLibraryService _styleLibraryService;
    private readonly Dictionary<string, Bitmap> _previewCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, SymbolDrawingIndex> _drawingIndexCache = new(StringComparer.OrdinalIgnoreCase);

    public SymbolPreviewService(ProjectStyleLibraryService styleLibraryService)
    {
        _styleLibraryService = styleLibraryService;
    }

    public DrawingImage GetPreview(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var drawingStamp = !string.IsNullOrWhiteSpace(context.SymbolStyleDrawingPath) && File.Exists(context.SymbolStyleDrawingPath)
            ? File.GetLastWriteTimeUtc(context.SymbolStyleDrawingPath).Ticks.ToString()
            : "nodwg";
        var cacheKey = $"{context.SymbolStyleDrawingPath}|{drawingStamp}|{item.UniqueKey}|{context.CurrentStandardToken}|{context.EffectiveRespectSupportedStandards}";
        if (_previewCache.TryGetValue(cacheKey, out var cached))
        {
            return cached;
        }

        var styleResolution = GetStyleResolution(item, context);
        var styleInfo = styleResolution.StyleInfo;

        var bitmap = item.Category == PaletteCategory.Lines
            ? TryCreateStyledLinePreview(item, styleInfo)
                ?? CreateLineFallbackPreview(item, styleInfo)
            : TryLoadResolvedBlockPreview(context, styleResolution.DisplayBlockName)
                ?? CreateFallbackPreview(item, styleInfo);

        _previewCache[cacheKey] = bitmap;
        return bitmap;
    }

    public void Dispose()
    {
        foreach (var bitmap in _previewCache.Values)
        {
            bitmap.Dispose();
        }

        foreach (var drawingIndex in _drawingIndexCache.Values)
        {
            drawingIndex.Dispose();
        }

        _previewCache.Clear();
        _drawingIndexCache.Clear();
    }

    public ProjectStyleResolution GetStyleResolution(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        return _styleLibraryService.Resolve(item, context);
    }

    private Bitmap? TryLoadResolvedBlockPreview(ProjectRuntimeContext context, string? blockName)
    {
        if (string.IsNullOrWhiteSpace(blockName))
        {
            return null;
        }

        var drawingPath = context.SymbolStyleDrawingPath;
        if (string.IsNullOrWhiteSpace(drawingPath) || !File.Exists(drawingPath))
        {
            return null;
        }

        var index = GetOrCreateDrawingIndex(drawingPath);
        if (!index.IsAvailable)
        {
            return null;
        }

        return index.TryGetPreview(blockName, out var preview)
            ? preview
            : null;
    }

    private Bitmap? TryCreateStyledLinePreview(ProjectPaletteItem item, ProjectStyleInfo? styleInfo)
    {
        return styleInfo is null ? null : CreateLinePreview(item, styleInfo);
    }


    private SymbolDrawingIndex GetOrCreateDrawingIndex(string drawingPath)
    {
        var cacheKey = $"{drawingPath}|{File.GetLastWriteTimeUtc(drawingPath).Ticks}";
        if (_drawingIndexCache.TryGetValue(cacheKey, out var cached))
        {
            return cached;
        }

        cached = BuildDrawingIndex(drawingPath);
        _drawingIndexCache[cacheKey] = cached;
        return cached;
    }

    private static SymbolDrawingIndex BuildDrawingIndex(string drawingPath)
    {
        if (!File.Exists(drawingPath))
        {
            return SymbolDrawingIndex.Unavailable;
        }

        var exactBitmaps = new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var database = new AcadDb.Database(false, true);
            database.ReadDwgFile(drawingPath, AcadDb.FileOpenMode.OpenForReadAndAllShare, false, null);

            var previousWorkingDatabase = AcadDb.HostApplicationServices.WorkingDatabase;
            try
            {
                AcadDb.HostApplicationServices.WorkingDatabase = database;

                using var transaction = database.TransactionManager.StartTransaction();
                var blockTable = (AcadDb.BlockTable)transaction.GetObject(database.BlockTableId, AcadDb.OpenMode.ForRead);

                foreach (AcadDb.ObjectId blockId in blockTable)
                {
                    var block = (AcadDb.BlockTableRecord)transaction.GetObject(blockId, AcadDb.OpenMode.ForRead);
                    if (block.IsAnonymous || block.IsLayout)
                    {
                        continue;
                    }

                    var blockName = block.Name?.Trim();
                    if (string.IsNullOrWhiteSpace(blockName))
                    {
                        continue;
                    }

                    var bitmap = TryCreateBlockBitmap(block, transaction, database);
                    if (bitmap is not null)
                    {
                        exactBitmaps[blockName] = bitmap;
                    }
                }
            }
            finally
            {
                try
                {
                    AcadDb.HostApplicationServices.WorkingDatabase = previousWorkingDatabase;
                }
                catch
                {
                }
            }
        }
        catch
        {
            foreach (var bitmap in exactBitmaps.Values)
            {
                bitmap.Dispose();
            }

            return SymbolDrawingIndex.Unavailable;
        }

        return new SymbolDrawingIndex(exactBitmaps);
    }

    private static Bitmap? TryCreateBlockBitmap(AcadDb.BlockTableRecord block, AcadDb.Transaction transaction, AcadDb.Database database)
    {
        if (block.HasPreviewIcon && block.PreviewIcon is not null)
        {
            return NormalizePreviewBitmap(block.PreviewIcon);
        }

        try
        {
            var rendered = RenderBlockPreview(block, transaction, database);
            if (rendered is not null)
            {
                return rendered;
            }
        }
        catch
        {
        }

        return null;
    }

    private static Bitmap NormalizePreviewBitmap(Bitmap source)
    {
        var bitmap = new Bitmap(PreviewSize, PreviewSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, PreviewSize - 1, PreviewSize - 1);

        const float padding = 6f;
        var scaleX = (PreviewSize - (padding * 2f)) / Math.Max(1f, source.Width);
        var scaleY = (PreviewSize - (padding * 2f)) / Math.Max(1f, source.Height);
        var scale = Math.Min(scaleX, scaleY);
        var drawWidth = source.Width * scale;
        var drawHeight = source.Height * scale;
        var offsetX = (PreviewSize - drawWidth) / 2f;
        var offsetY = (PreviewSize - drawHeight) / 2f;

        graphics.DrawImage(source, offsetX, offsetY, drawWidth, drawHeight);
        return bitmap;
    }

    private static Bitmap? RenderBlockPreview(AcadDb.BlockTableRecord block, AcadDb.Transaction transaction, AcadDb.Database database)
    {
        var strokes = new List<RenderStroke>();
        var fills = new List<RenderFill>();
        var transformStack = new List<AcadGeom.Matrix3d>();
        var visitedBlocks = new HashSet<AcadDb.ObjectId>();

        CollectRenderData(block, transaction, database, transformStack, visitedBlocks, strokes, fills, depth: 0);
        if (strokes.Count == 0 && fills.Count == 0)
        {
            return null;
        }

        var bounds = ComputeBounds(strokes, fills);
        if (!bounds.HasValue)
        {
            return null;
        }

        return DrawPreview(bounds.Value, strokes, fills);
    }

    private static void CollectRenderData(
        AcadDb.BlockTableRecord block,
        AcadDb.Transaction transaction,
        AcadDb.Database database,
        IList<AcadGeom.Matrix3d> transformStack,
        ISet<AcadDb.ObjectId> visitedBlocks,
        ICollection<RenderStroke> strokes,
        ICollection<RenderFill> fills,
        int depth)
    {
        if (depth > 8 || !visitedBlocks.Add(block.ObjectId))
        {
            return;
        }

        foreach (AcadDb.ObjectId entityId in block)
        {
            AcadDb.Entity? entity;
            try
            {
                entity = transaction.GetObject(entityId, AcadDb.OpenMode.ForRead, false) as AcadDb.Entity;
            }
            catch
            {
                continue;
            }

            if (entity is null || entity.Visible == false)
            {
                continue;
            }

            switch (entity)
            {
                case AcadDb.BlockReference blockReference:
                    TryCollectNestedBlock(blockReference, transaction, database, transformStack, visitedBlocks, strokes, fills, depth);
                    break;

                case AcadDb.Solid solid:
                    TryCollectSolid(solid, transaction, database, transformStack, strokes, fills);
                    break;


                case AcadDb.Curve curve:
                    TryCollectCurve(curve, transaction, database, transformStack, strokes);
                    break;
            }
        }

        visitedBlocks.Remove(block.ObjectId);
    }

    private static void TryCollectNestedBlock(
        AcadDb.BlockReference blockReference,
        AcadDb.Transaction transaction,
        AcadDb.Database database,
        IList<AcadGeom.Matrix3d> transformStack,
        ISet<AcadDb.ObjectId> visitedBlocks,
        ICollection<RenderStroke> strokes,
        ICollection<RenderFill> fills,
        int depth)
    {
        AcadDb.BlockTableRecord? nestedBlock;
        try
        {
            nestedBlock = transaction.GetObject(blockReference.BlockTableRecord, AcadDb.OpenMode.ForRead, false) as AcadDb.BlockTableRecord;
        }
        catch
        {
            return;
        }

        if (nestedBlock is null)
        {
            return;
        }

        transformStack.Add(blockReference.BlockTransform);
        CollectRenderData(nestedBlock, transaction, database, transformStack, visitedBlocks, strokes, fills, depth + 1);
        transformStack.RemoveAt(transformStack.Count - 1);
    }

    private static void TryCollectSolid(
        AcadDb.Solid solid,
        AcadDb.Transaction transaction,
        AcadDb.Database database,
        IList<AcadGeom.Matrix3d> transformStack,
        ICollection<RenderStroke> strokes,
        ICollection<RenderFill> fills)
    {
        try
        {
            var points = new List<AcadGeom.Point3d>
            {
                ApplyTransforms(solid.GetPointAt(0), transformStack),
                ApplyTransforms(solid.GetPointAt(1), transformStack),
                ApplyTransforms(solid.GetPointAt(2), transformStack),
                ApplyTransforms(solid.GetPointAt(3), transformStack)
            };

            var distinctPoints = DeduplicateSequentialPoints(points);
            if (distinctPoints.Count >= 3)
            {
                var color = ResolveEntityColor(solid, transaction, database);
                fills.Add(new RenderFill(distinctPoints, Color.FromArgb(140, color)));
                strokes.Add(new RenderStroke(ClosePoints(distinctPoints), color, ResolveEntityStrokeWidth(solid)));
            }
        }
        catch
        {
        }
    }


    private static void TryCollectCurve(
        AcadDb.Curve curve,
        AcadDb.Transaction transaction,
        AcadDb.Database database,
        IList<AcadGeom.Matrix3d> transformStack,
        ICollection<RenderStroke> strokes)
    {
        try
        {
            var pointCount = EstimateSampleCount(curve);
            if (pointCount < 2)
            {
                return;
            }

            var start = curve.StartParam;
            var end = curve.EndParam;
            var points = new List<AcadGeom.Point3d>(pointCount + 1);

            for (var i = 0; i <= pointCount; i++)
            {
                var param = start + ((end - start) * i / pointCount);
                var point = curve.GetPointAtParameter(param);
                points.Add(ApplyTransforms(point, transformStack));
            }

            points = DeduplicateSequentialPoints(points);
            if (points.Count < 2)
            {
                return;
            }

            if (IsClosedCurve(curve))
            {
                points = ClosePoints(points);
            }

            strokes.Add(new RenderStroke(points, ResolveEntityColor(curve, transaction, database), ResolveEntityStrokeWidth(curve)));
        }
        catch
        {
        }
    }

    private static int EstimateSampleCount(AcadDb.Curve curve)
    {
        return curve switch
        {
            AcadDb.Line => 2,
            AcadDb.Arc arc => Math.Max(12, (int)Math.Ceiling(Math.Abs(arc.EndAngle - arc.StartAngle) / (Math.PI / 12d))),
            AcadDb.Circle => 40,
            AcadDb.Ellipse => 40,
            AcadDb.Polyline polyline => Math.Max(4, polyline.NumberOfVertices * 6),
            AcadDb.Polyline2d => 24,
            AcadDb.Polyline3d => 24,
            AcadDb.Spline => 40,
            _ => 24
        };
    }

    private static bool IsClosedCurve(AcadDb.Curve curve)
    {
        if (curve is AcadDb.Circle)
        {
            return true;
        }

        var property = curve.GetType().GetProperty("Closed", BindingFlags.Instance | BindingFlags.Public);
        if (property is not null)
        {
            try
            {
                if (property.GetValue(curve) is bool isClosed)
                {
                    return isClosed;
                }
            }
            catch
            {
            }
        }

        return false;
    }

    private static List<AcadGeom.Point3d> ClosePoints(IReadOnlyList<AcadGeom.Point3d> points)
    {
        var result = points.ToList();
        if (result.Count > 1 && result[0].DistanceTo(result[^1]) > 1e-6)
        {
            result.Add(result[0]);
        }

        return result;
    }

    private static List<AcadGeom.Point3d> DeduplicateSequentialPoints(IReadOnlyList<AcadGeom.Point3d> points)
    {
        var result = new List<AcadGeom.Point3d>(points.Count);
        foreach (var point in points)
        {
            if (result.Count == 0 || result[^1].DistanceTo(point) > 1e-6)
            {
                result.Add(point);
            }
        }

        return result;
    }

    private static AcadGeom.Point3d ApplyTransforms(AcadGeom.Point3d point, IList<AcadGeom.Matrix3d> transformStack)
    {
        var result = point;
        for (var i = transformStack.Count - 1; i >= 0; i--)
        {
            result = result.TransformBy(transformStack[i]);
        }

        return result;
    }

    private static RectangleF? ComputeBounds(IEnumerable<RenderStroke> strokes, IEnumerable<RenderFill> fills)
    {
        var hasPoint = false;
        var minX = double.MaxValue;
        var minY = double.MaxValue;
        var maxX = double.MinValue;
        var maxY = double.MinValue;

        foreach (var point in strokes.SelectMany(stroke => stroke.Points).Concat(fills.SelectMany(fill => fill.Points)))
        {
            hasPoint = true;
            minX = Math.Min(minX, point.X);
            minY = Math.Min(minY, point.Y);
            maxX = Math.Max(maxX, point.X);
            maxY = Math.Max(maxY, point.Y);
        }

        if (!hasPoint)
        {
            return null;
        }

        var width = Math.Max(1e-6, maxX - minX);
        var height = Math.Max(1e-6, maxY - minY);
        return new RectangleF((float)minX, (float)minY, (float)width, (float)height);
    }

    private static Bitmap DrawPreview(RectangleF modelBounds, IReadOnlyCollection<RenderStroke> strokes, IReadOnlyCollection<RenderFill> fills)
    {
        var bitmap = new Bitmap(PreviewSize, PreviewSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, PreviewSize - 1, PreviewSize - 1);

        var workingBounds = InflateBounds(modelBounds, 0.08f, minimumInflate: 0.6f);
        const float padding = 8f;
        var scaleX = (PreviewSize - (padding * 2f)) / Math.Max(0.01f, workingBounds.Width);
        var scaleY = (PreviewSize - (padding * 2f)) / Math.Max(0.01f, workingBounds.Height);
        var scale = Math.Min(scaleX, scaleY);
        var drawWidth = workingBounds.Width * scale;
        var drawHeight = workingBounds.Height * scale;
        var offsetX = (PreviewSize - drawWidth) / 2f;
        var offsetY = (PreviewSize - drawHeight) / 2f;

        PointF Map(AcadGeom.Point3d point)
        {
            var x = offsetX + (float)((point.X - workingBounds.Left) * scale);
            var y = PreviewSize - offsetY - (float)((point.Y - workingBounds.Top) * scale);
            return new PointF(x, y);
        }

        foreach (var fill in fills)
        {
            var points = fill.Points.Select(Map).ToArray();
            if (points.Length >= 3)
            {
                using var brush = new SolidBrush(fill.Color);
                graphics.FillPolygon(brush, points, FillMode.Winding);
            }
        }

        foreach (var stroke in strokes)
        {
            var points = stroke.Points.Select(Map).ToArray();
            if (points.Length < 2)
            {
                continue;
            }

            var penWidth = Math.Clamp(stroke.Width * Math.Max(1f, PreviewSize / 72f), 1.6f, 4.6f);
            using var pen = new Pen(stroke.Color, penWidth)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };
            graphics.DrawLines(pen, points);
        }

        return bitmap;
    }

    private static RectangleF InflateBounds(RectangleF bounds, float scaleFactor, float minimumInflate)
    {
        var inflateX = Math.Max(minimumInflate, bounds.Width * scaleFactor);
        var inflateY = Math.Max(minimumInflate, bounds.Height * scaleFactor);
        bounds.Inflate(inflateX, inflateY);
        return bounds;
    }

    private static Color ResolveEntityColor(AcadDb.Entity entity, AcadDb.Transaction transaction, AcadDb.Database database)
    {
        try
        {
            var color = entity.Color;
            if (color is not null)
            {
                try
                {
                    return color.ColorValue;
                }
                catch
                {
                }

                if (color.ColorMethod == AcadColors.ColorMethod.ByLayer && !entity.LayerId.IsNull)
                {
                    if (transaction.GetObject(entity.LayerId, AcadDb.OpenMode.ForRead, false) is AcadDb.LayerTableRecord layer)
                    {
                        try
                        {
                            return layer.Color.ColorValue;
                        }
                        catch
                        {
                        }
                    }
                }
            }
        }
        catch
        {
        }

        return Color.Black;
    }

    private static float ResolveEntityStrokeWidth(AcadDb.Entity entity)
    {
        try
        {
            var lineWeight = (int)entity.LineWeight;
            if (lineWeight > 0)
            {
                return Math.Clamp(lineWeight / 26f, 1.4f, 3.6f);
            }
        }
        catch
        {
        }

        return 1.6f;
    }

    private static IEnumerable<string> BuildBlockAlternativeKeys(string value)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Add(value);
        Add(SplitPascalCase(value));
        Add(RemoveTrailingWord(value, "Style"));
        Add(RemoveTrailingWord(value, "Block"));
        Add($"{value} Style");
        Add($"{value} Block");
        Add(value.Replace("-", " ", StringComparison.Ordinal));
        Add(value.Replace("/", " ", StringComparison.Ordinal));
        Add(value.Replace("_", " ", StringComparison.Ordinal));
        Add(SearchText.Compact(value));

        foreach (var item in seen)
        {
            yield return item;
        }

        void Add(string? candidate)
        {
            var trimmed = candidate?.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
            {
                seen.Add(trimmed);
            }
        }
    }

    private static Bitmap CreateLinePreview(ProjectPaletteItem item, ProjectStyleInfo styleInfo)
    {
        var bitmap = new Bitmap(PreviewSize, PreviewSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, PreviewSize - 1, PreviewSize - 1);

        using var linePen = new Pen(ResolveStyleColor(styleInfo, item), ResolveStyleStrokeWidth(styleInfo))
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };
        ApplyDashPattern(linePen, styleInfo, item);

        if (SearchText.ContainsPhrase(SearchText.Normalize(item.LineStyleName ?? item.VisualName), "jacketed"))
        {
            graphics.DrawLine(linePen, 8, 26, PreviewSize - 8, 26);
            graphics.DrawLine(linePen, 8, 38, PreviewSize - 8, 38);
        }
        else
        {
            graphics.DrawLine(linePen, 8, PreviewSize / 2f, PreviewSize - 8, PreviewSize / 2f);
        }

        return bitmap;
    }

    private static Bitmap CreateLineFallbackPreview(ProjectPaletteItem item, ProjectStyleInfo? styleInfo)
    {
        if (styleInfo is not null)
        {
            return CreateLinePreview(item, styleInfo);
        }

        var bitmap = new Bitmap(PreviewSize, PreviewSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, PreviewSize - 1, PreviewSize - 1);

        var styleText = SearchText.Normalize(item.LineStyleName ?? item.VisualName);
        using var linePen = new Pen(ResolveLineColor(styleText), styleText.Contains("secondary", StringComparison.Ordinal) ? 2f : 2.5f)
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round
        };

        if (styleText.Contains("signal", StringComparison.Ordinal) || styleText.Contains("electrical", StringComparison.Ordinal) || styleText.Contains("pneumatic", StringComparison.Ordinal))
        {
            linePen.DashPattern = [6f, 3f];
        }
        else if (styleText.Contains("capillary", StringComparison.Ordinal) || styleText.Contains("hydraulic", StringComparison.Ordinal) || styleText.Contains("mechanical", StringComparison.Ordinal))
        {
            linePen.DashPattern = [1f, 2f];
        }
        else if (styleText.Contains("existing", StringComparison.Ordinal))
        {
            linePen.DashPattern = [9f, 3f, 2f, 3f];
        }

        if (styleText.Contains("jacketed", StringComparison.Ordinal))
        {
            graphics.DrawLine(linePen, 8, 26, PreviewSize - 8, 26);
            graphics.DrawLine(linePen, 8, 38, PreviewSize - 8, 38);
        }
        else
        {
            graphics.DrawLine(linePen, 8, PreviewSize / 2f, PreviewSize - 8, PreviewSize / 2f);
        }

        return bitmap;
    }

    private static Bitmap CreateFallbackPreview(ProjectPaletteItem item, ProjectStyleInfo? styleInfo)
    {
        var bitmap = new Bitmap(PreviewSize, PreviewSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, PreviewSize - 1, PreviewSize - 1);

        var accentColor = styleInfo?.ArgbColor is int argbColor
            ? Color.FromArgb(argbColor)
            : Color.Black;

        using var pen = new Pen(accentColor, 2f)
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };

        switch (item.Category)
        {
            case PaletteCategory.Equipment:
                graphics.DrawEllipse(pen, 12, 12, PreviewSize - 24, PreviewSize - 24);
                break;

            case PaletteCategory.Valves:
                graphics.DrawLine(pen, 8, PreviewSize / 2f, 22, PreviewSize / 2f);
                graphics.DrawLine(pen, PreviewSize - 22, PreviewSize / 2f, PreviewSize - 8, PreviewSize / 2f);
                graphics.DrawPolygon(pen, [new PointF(22, PreviewSize / 2f), new PointF(32, 18), new PointF(42, PreviewSize / 2f), new PointF(32, 46)]);
                break;

            case PaletteCategory.Fittings:
            case PaletteCategory.Reducers:
                graphics.DrawLine(pen, 8, PreviewSize / 2f, PreviewSize - 8, PreviewSize / 2f);
                graphics.DrawPolygon(pen, [new PointF(22, 18), new PointF(42, PreviewSize / 2f), new PointF(22, 46)]);
                break;

            case PaletteCategory.Instrumentation:
                graphics.DrawEllipse(pen, 16, 16, PreviewSize - 32, PreviewSize - 32);
                graphics.DrawLine(pen, PreviewSize / 2f, 8, PreviewSize / 2f, 16);
                graphics.DrawLine(pen, PreviewSize / 2f, PreviewSize - 16, PreviewSize / 2f, PreviewSize - 8);
                break;

            case PaletteCategory.Nozzles:
                graphics.DrawLine(pen, 12, PreviewSize / 2f, 36, PreviewSize / 2f);
                graphics.DrawEllipse(pen, 36, 24, 16, 16);
                break;

            default:
                graphics.DrawRectangle(pen, 14, 14, PreviewSize - 28, PreviewSize - 28);
                graphics.DrawLine(pen, 14, PreviewSize / 2f, PreviewSize - 14, PreviewSize / 2f);
                break;
        }

        return bitmap;
    }

    private static Color ResolveStyleColor(ProjectStyleInfo styleInfo, ProjectPaletteItem item)
    {
        if (styleInfo.ArgbColor is int argb)
        {
            return Color.FromArgb(argb);
        }

        return ResolveLineColor(SearchText.Normalize(styleInfo.LinetypeName ?? styleInfo.StyleName ?? item.VisualName));
    }

    private static float ResolveStyleStrokeWidth(ProjectStyleInfo styleInfo)
    {
        if (styleInfo.LineWeight is int lineWeight && lineWeight > 0)
        {
            return Math.Clamp(lineWeight / 30f, 1.3f, 3.5f);
        }

        return 2.2f;
    }

    private static void ApplyDashPattern(Pen pen, ProjectStyleInfo styleInfo, ProjectPaletteItem item)
    {
        var lookupText = SearchText.Normalize($"{styleInfo.LinetypeName} {styleInfo.StyleName} {item.LineStyleName} {item.VisualName}");
        if (lookupText.Contains("dash dot", StringComparison.Ordinal) || lookupText.Contains("center", StringComparison.Ordinal))
        {
            pen.DashPattern = [8f, 3f, 2f, 3f];
        }
        else if (lookupText.Contains("hidden", StringComparison.Ordinal) || lookupText.Contains("dashed", StringComparison.Ordinal) || lookupText.Contains("signal", StringComparison.Ordinal))
        {
            pen.DashPattern = [6f, 3f];
        }
        else if (lookupText.Contains("capillary", StringComparison.Ordinal) || lookupText.Contains("dot", StringComparison.Ordinal))
        {
            pen.DashPattern = [1f, 2f];
        }
    }

    private static Color ResolveLineColor(string styleText)
    {
        if (styleText.Contains("existing", StringComparison.Ordinal))
        {
            return Color.DimGray;
        }

        if (styleText.Contains("secondary", StringComparison.Ordinal))
        {
            return Color.DarkSlateGray;
        }

        if (styleText.Contains("signal", StringComparison.Ordinal)
            || styleText.Contains("electrical", StringComparison.Ordinal)
            || styleText.Contains("electromagnetic", StringComparison.Ordinal))
        {
            return Color.SteelBlue;
        }

        return Color.Black;
    }

    private static string Abbreviate(string value)
    {
        var tokens = value.Split([' ', '-', '_'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return tokens.Length switch
        {
            0 => value.Length <= 4 ? value.ToUpperInvariant() : value[..4].ToUpperInvariant(),
            1 => tokens[0].Length <= 4 ? tokens[0].ToUpperInvariant() : tokens[0][..4].ToUpperInvariant(),
            _ => string.Concat(tokens.Take(3).Select(token => char.ToUpperInvariant(token[0])))
        };
    }

    private static string RemoveTrailingWord(string value, string trailingWord)
    {
        var trimmed = value.Trim();
        while (trimmed.EndsWith($" {trailingWord}", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith($"_{trailingWord}", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith(trailingWord, StringComparison.OrdinalIgnoreCase))
        {
            if (trimmed.EndsWith($" {trailingWord}", StringComparison.OrdinalIgnoreCase)
                || trimmed.EndsWith($"_{trailingWord}", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed[..^(trailingWord.Length + 1)].TrimEnd('_', ' ');
                continue;
            }

            var boundaryIndex = trimmed.Length - trailingWord.Length;
            var hasBoundary = boundaryIndex == 0 || !char.IsLetterOrDigit(trimmed[boundaryIndex - 1]);
            if (!hasBoundary)
            {
                break;
            }

            trimmed = trimmed[..boundaryIndex].TrimEnd('_', ' ');
        }

        return string.IsNullOrWhiteSpace(trimmed) ? value : trimmed;
    }

    private static string SplitPascalCase(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var chars = new List<char>(value.Length * 2);
        for (var i = 0; i < value.Length; i++)
        {
            var ch = value[i];
            if (i > 0 && char.IsUpper(ch) && (char.IsLower(value[i - 1]) || (i + 1 < value.Length && char.IsLower(value[i + 1]))))
            {
                chars.Add(' ');
            }

            chars.Add(ch);
        }

        return new string(chars.ToArray());
    }

    private sealed record RenderStroke(IReadOnlyList<AcadGeom.Point3d> Points, Color Color, float Width);

    private sealed record RenderFill(IReadOnlyList<AcadGeom.Point3d> Points, Color Color);

    private sealed class SymbolDrawingIndex : IDisposable
    {
        public static SymbolDrawingIndex Unavailable { get; } = new(
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase));

        private readonly IReadOnlyDictionary<string, Bitmap> _exactBitmaps;

        public SymbolDrawingIndex(IReadOnlyDictionary<string, Bitmap> exactBitmaps)
        {
            _exactBitmaps = exactBitmaps;
        }

        public bool IsAvailable => _exactBitmaps.Count > 0;

        public bool TryGetPreview(string candidateName, out Bitmap preview)
        {
            preview = null!;
            if (string.IsNullOrWhiteSpace(candidateName))
            {
                return false;
            }

            if (_exactBitmaps.TryGetValue(candidateName.Trim(), out var exactBitmap))
            {
                preview = new Bitmap(exactBitmap);
                return true;
            }

            return false;
        }

        public void Dispose()
        {
            foreach (var bitmap in _exactBitmaps.Values)
            {
                bitmap.Dispose();
            }
        }
    }

}
