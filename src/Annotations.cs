using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private enum PointerTool
        {
            Arrow,
            Laser,
            Pen,
            Highlighter,
            Eraser
        }

        private sealed class AnnotationStroke
        {
            public readonly List<PointF> Points = new List<PointF>();
            public Color Color;
            public float Width;
        }

        private readonly Dictionary<int, List<AnnotationStroke>> slideAnnotations =
            new Dictionary<int, List<AnnotationStroke>>();

        private PointerTool pointerTool = PointerTool.Arrow;
        private AnnotationStroke activeStroke;
        private bool drawingAnnotation;
        private Point laserPoint;
        private bool laserVisible;
        private bool presentationToolsInitialized;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            if (presentationToolsInitialized)
                return;

            presentationToolsInitialized = true;

            viewer.Paint += OnViewerPaintPresentationTools;
            viewer.MouseDown += OnViewerPresentationMouseDown;
            viewer.MouseMove += OnViewerPresentationMouseMove;
            viewer.MouseUp += OnViewerPresentationMouseUp;
            viewer.MouseLeave += delegate
            {
                laserVisible = false;
                viewer.Invalidate();
            };

            AddPresentationPointerMenu();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (internalSlideShowMode)
            {
                if (keyData == (Keys.Control | Keys.A))
                {
                    SetPointerTool(PointerTool.Arrow);
                    return true;
                }

                if (keyData == (Keys.Control | Keys.L))
                {
                    SetPointerTool(PointerTool.Laser);
                    return true;
                }

                if (keyData == (Keys.Control | Keys.P))
                {
                    SetPointerTool(PointerTool.Pen);
                    return true;
                }

                if (keyData == (Keys.Control | Keys.H))
                {
                    SetPointerTool(PointerTool.Highlighter);
                    return true;
                }

                if (keyData == (Keys.Control | Keys.Shift | Keys.E))
                {
                    ClearCurrentSlideAnnotations();
                    return true;
                }

                if (keyData == (Keys.Control | Keys.E))
                {
                    SetPointerTool(PointerTool.Eraser);
                    return true;
                }
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void AddPresentationPointerMenu()
        {
            slideshowMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem pointerMenu =
                new ToolStripMenuItem("Pointer Options");

            pointerMenu.DropDownItems.Add(
                "Arrow",
                null,
                delegate { SetPointerTool(PointerTool.Arrow); });

            pointerMenu.DropDownItems.Add(
                "Laser Pointer",
                null,
                delegate { SetPointerTool(PointerTool.Laser); });

            pointerMenu.DropDownItems.Add(
                "Pen",
                null,
                delegate { SetPointerTool(PointerTool.Pen); });

            pointerMenu.DropDownItems.Add(
                "Highlighter",
                null,
                delegate { SetPointerTool(PointerTool.Highlighter); });

            pointerMenu.DropDownItems.Add(
                "Eraser",
                null,
                delegate { SetPointerTool(PointerTool.Eraser); });

            pointerMenu.DropDownItems.Add(new ToolStripSeparator());

            pointerMenu.DropDownItems.Add(
                "Erase Ink on Slide",
                null,
                delegate { ClearCurrentSlideAnnotations(); });

            pointerMenu.DropDownItems.Add(
                "Erase All Ink",
                null,
                delegate { ClearAllAnnotations(); });

            slideshowMenu.Items.Add(pointerMenu);
        }

        private void SetPointerTool(PointerTool tool)
        {
            pointerTool = tool;
            drawingAnnotation = false;
            activeStroke = null;
            laserVisible = false;

            viewer.Cursor =
                tool == PointerTool.Arrow
                    ? Cursors.Default
                    : Cursors.Cross;

            viewer.Invalidate();
        }

        private List<AnnotationStroke> GetSlideAnnotations(
            int slideIndex,
            bool create)
        {
            if (slideIndex < 0)
                return null;

            List<AnnotationStroke> strokes;

            if (slideAnnotations.TryGetValue(slideIndex, out strokes))
                return strokes;

            if (!create)
                return null;

            strokes = new List<AnnotationStroke>();
            slideAnnotations[slideIndex] = strokes;
            return strokes;
        }

        private PointF NormalizeViewerPoint(Point point)
        {
            float width = Math.Max(1, viewer.ClientSize.Width);
            float height = Math.Max(1, viewer.ClientSize.Height);

            return new PointF(
                Math.Max(0f, Math.Min(1f, point.X / width)),
                Math.Max(0f, Math.Min(1f, point.Y / height)));
        }

        private PointF DenormalizeViewerPoint(PointF point)
        {
            return new PointF(
                point.X * viewer.ClientSize.Width,
                point.Y * viewer.ClientSize.Height);
        }

        private void OnViewerPresentationMouseDown(
            object sender,
            MouseEventArgs e)
        {
            if (!internalSlideShowMode ||
                e.Button != MouseButtons.Left)
            {
                return;
            }

            if (pointerTool == PointerTool.Pen ||
                pointerTool == PointerTool.Highlighter)
            {
                AnnotationStroke stroke = new AnnotationStroke();

                stroke.Color =
                    pointerTool == PointerTool.Highlighter
                        ? Color.FromArgb(110, 255, 235, 59)
                        : Color.FromArgb(235, 220, 40, 40);

                stroke.Width =
                    pointerTool == PointerTool.Highlighter
                        ? 14f
                        : 3.5f;

                stroke.Points.Add(
                    NormalizeViewerPoint(e.Location));

                List<AnnotationStroke> strokes =
                    GetSlideAnnotations(currentIndex, true);

                if (strokes != null)
                    strokes.Add(stroke);

                activeStroke = stroke;
                drawingAnnotation = true;
                viewer.Invalidate();
            }
            else if (pointerTool == PointerTool.Eraser)
            {
                EraseAnnotationAt(e.Location);
            }
        }

        private void OnViewerPresentationMouseMove(
            object sender,
            MouseEventArgs e)
        {
            if (!internalSlideShowMode)
                return;

            if (pointerTool == PointerTool.Laser)
            {
                laserPoint = e.Location;
                laserVisible = true;
                viewer.Cursor = Cursors.Cross;
                viewer.Invalidate();
                return;
            }

            if ((pointerTool == PointerTool.Pen ||
                 pointerTool == PointerTool.Highlighter) &&
                drawingAnnotation &&
                activeStroke != null)
            {
                activeStroke.Points.Add(
                    NormalizeViewerPoint(e.Location));

                viewer.Invalidate();
            }
        }

        private void OnViewerPresentationMouseUp(
            object sender,
            MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            drawingAnnotation = false;
            activeStroke = null;
        }

        private void OnViewerPaintPresentationTools(
            object sender,
            PaintEventArgs e)
        {
            if (!internalSlideShowMode || currentIndex < 0)
                return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            List<AnnotationStroke> strokes =
                GetSlideAnnotations(currentIndex, false);

            if (strokes != null)
            {
                for (int i = 0; i < strokes.Count; i++)
                    DrawAnnotationStroke(e.Graphics, strokes[i]);
            }

            if (pointerTool == PointerTool.Laser &&
                laserVisible)
            {
                const float outer = 18f;
                const float inner = 8f;

                using (Brush glow =
                    new SolidBrush(Color.FromArgb(90, 255, 0, 0)))
                {
                    e.Graphics.FillEllipse(
                        glow,
                        laserPoint.X - outer / 2f,
                        laserPoint.Y - outer / 2f,
                        outer,
                        outer);
                }

                using (Brush dot =
                    new SolidBrush(Color.FromArgb(245, 255, 35, 35)))
                {
                    e.Graphics.FillEllipse(
                        dot,
                        laserPoint.X - inner / 2f,
                        laserPoint.Y - inner / 2f,
                        inner,
                        inner);
                }
            }
        }

        private void DrawAnnotationStroke(
            Graphics graphics,
            AnnotationStroke stroke)
        {
            if (stroke == null || stroke.Points.Count == 0)
                return;

            using (Pen pen =
                new Pen(stroke.Color, stroke.Width))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;

                if (stroke.Points.Count == 1)
                {
                    PointF point =
                        DenormalizeViewerPoint(stroke.Points[0]);

                    using (Brush brush =
                        new SolidBrush(stroke.Color))
                    {
                        float size = Math.Max(2f, stroke.Width);

                        graphics.FillEllipse(
                            brush,
                            point.X - size / 2f,
                            point.Y - size / 2f,
                            size,
                            size);
                    }

                    return;
                }

                PointF[] points =
                    new PointF[stroke.Points.Count];

                for (int i = 0; i < stroke.Points.Count; i++)
                    points[i] =
                        DenormalizeViewerPoint(stroke.Points[i]);

                graphics.DrawLines(pen, points);
            }
        }

        private void EraseAnnotationAt(Point location)
        {
            List<AnnotationStroke> strokes =
                GetSlideAnnotations(currentIndex, false);

            if (strokes == null || strokes.Count == 0)
                return;

            float bestDistance = float.MaxValue;
            int bestIndex = -1;

            for (int i = 0; i < strokes.Count; i++)
            {
                AnnotationStroke stroke = strokes[i];

                for (int p = 0; p < stroke.Points.Count; p++)
                {
                    PointF point =
                        DenormalizeViewerPoint(stroke.Points[p]);

                    float dx = point.X - location.X;
                    float dy = point.Y - location.Y;
                    float distance = dx * dx + dy * dy;

                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestIndex = i;
                    }
                }
            }

            const float eraserRadius = 28f;

            if (bestIndex >= 0 &&
                bestDistance <= eraserRadius * eraserRadius)
            {
                strokes.RemoveAt(bestIndex);
                viewer.Invalidate();
            }
        }

        private void ClearCurrentSlideAnnotations()
        {
            if (currentIndex >= 0)
                slideAnnotations.Remove(currentIndex);

            viewer.Invalidate();
        }

        private void ClearAllAnnotations()
        {
            slideAnnotations.Clear();
            viewer.Invalidate();
        }
    }
}
