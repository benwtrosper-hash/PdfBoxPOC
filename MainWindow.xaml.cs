using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

// Aliases to avoid WPF collisions
using Drawing = System.Drawing;
using IOPath = System.IO.Path;

// Alias the two PdfDocument types to avoid ambiguity
using PdfiumDoc = PdfiumViewer.PdfDocument;
using PdfiumFlags = PdfiumViewer.PdfRenderFlags;

using PigDoc = UglyToad.PdfPig.PdfDocument;

namespace PdfBoxPOC
{
    public partial class MainWindow : Window
    {
        // =========================
        // PDF / Render
        // =========================
        private PdfiumDoc? _pdfViewerDoc;
        private string? _currentPdfPath;
        private int _pageIndex = 0; // 0-based
        private Drawing.SizeF _pageSizePoints;
        private int _renderDpi = 150;

        // =========================
        // Selection (transient)
        // =========================
        private bool _isDrawing;
        private Point _start;
        private System.Windows.Shapes.Rectangle? _selectionRectShape;

        private PdfRect? _lastSelectionPdfRect; // top-left origin points
        private string _lastSelectionText = "";

        // =========================
        // Editing mode
        // =========================
        private enum EditMode { Normal, RedrawRectForItem }
        private EditMode _editMode = EditMode.Normal;
        private Guid _pendingEditItemId = Guid.Empty;

        // =========================
        // Overlay visuals
        // =========================
        private sealed class OverlayVisual
        {
            public Guid Id;
            public System.Windows.Shapes.Rectangle Rect = null!;
            public TextBlock Label = null!;
            public ItemKind Kind;
        }

        private readonly Dictionary<Guid, OverlayVisual> _visuals = new();

        // =========================
        // Undo / Redo
        // =========================
        private interface IUndoableCommand
        {
            string Name { get; }
            void Do();
            void Undo();
        }

        private readonly Stack<IUndoableCommand> _undo = new();
        private readonly Stack<IUndoableCommand> _redo = new();

        private void PushAndDo(IUndoableCommand cmd)
        {
            cmd.Do();
            _undo.Push(cmd);
            _redo.Clear();
            ViewModel.RectStatus = $"OK: {cmd.Name}";
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (_undo.Count == 0) { ViewModel.RectStatus = "Nothing to undo."; return; }
            var cmd = _undo.Pop();
            cmd.Undo();
            _redo.Push(cmd);
            ViewModel.RectStatus = $"Undo: {cmd.Name}";
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            if (_redo.Count == 0) { ViewModel.RectStatus = "Nothing to redo."; return; }
            var cmd = _redo.Pop();
            cmd.Do();
            _undo.Push(cmd);
            ViewModel.RectStatus = $"Redo: {cmd.Name}";
        }

        // =========================
        // ViewModel
        // =========================
        public Vm ViewModel { get; } = new Vm();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = ViewModel;

            ViewModel.RectStatus = "Open a PDF, then drag boxes to map fields/tables.";
            ViewModel.ScreenRectText = "ScreenRect: (none)";
            ViewModel.PdfRectText = "PdfRect: (none)";

            // Keep AllItems list current
            ViewModel.Fields.CollectionChanged += (_, __) => RebuildAllItems();
            ViewModel.Tables.CollectionChanged += (_, __) => RebuildAllItems();

            // Prop change sync for selection (guarded by VM setters + SelectDesignerItem guard)
            ViewModel.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(Vm.SelectedField) && ViewModel.SelectedField != null)
                    SelectDesignerItem(ViewModel.SelectedField);
                if (e.PropertyName == nameof(Vm.SelectedTable) && ViewModel.SelectedTable != null)
                    SelectDesignerItem(ViewModel.SelectedTable);
                if (e.PropertyName == nameof(Vm.SelectedColumn) && ViewModel.SelectedColumn != null)
                    SelectDesignerItem(ViewModel.SelectedColumn);
                if (e.PropertyName == nameof(Vm.SelectedItem) && ViewModel.SelectedItem != null)
                    SelectDesignerItem(ViewModel.SelectedItem);
            };
        }

        // =========================
        // PDF load / render
        // =========================
        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "Select a PDF"
            };
            if (dlg.ShowDialog() != true) return;
            LoadPdf(dlg.FileName);
        }

        private void LoadPdf(string path)
        {
            ClearSelectionOnly();

            _pdfViewerDoc?.Dispose();
            _pdfViewerDoc = PdfiumDoc.Load(path);
            _currentPdfPath = path;

            _pageIndex = 0;
            _pageSizePoints = _pdfViewerDoc.PageSizes[_pageIndex];

            using (var bmp = RenderPageToBitmap(_pdfViewerDoc, _pageIndex, _renderDpi))
            {
                PdfImage.Source = ToBitmapSource(bmp);

                PersistedLayer.Width = PdfImage.Source.Width;
                PersistedLayer.Height = PdfImage.Source.Height;

                SelectionLayer.Width = PdfImage.Source.Width;
                SelectionLayer.Height = PdfImage.Source.Height;
            }

            ViewModel.NotesText =
                $"Loaded: {IOPath.GetFileName(path)} | PageSize(points): {_pageSizePoints.Width:0.##} x {_pageSizePoints.Height:0.##}\n" +
                "Tip: Add Field / Add Table → Set Table Region → Add Column.\n" +
                "Click overlays to select; Edit Selected can rename or redraw.";
            ViewModel.RectStatus = "Ready.";

            RedrawPersistedOverlays();
        }

        private static Drawing.Bitmap RenderPageToBitmap(PdfiumDoc pdf, int pageIndex, int dpi)
        {
            var sizePts = pdf.PageSizes[pageIndex];
            int pxW = (int)Math.Round((sizePts.Width / 72f) * dpi);
            int pxH = (int)Math.Round((sizePts.Height / 72f) * dpi);

            using Drawing.Image img = pdf.Render(pageIndex, pxW, pxH, dpi, dpi, PdfiumFlags.Annotations);
            return new Drawing.Bitmap(img);
        }

        private static BitmapSource ToBitmapSource(Drawing.Bitmap bitmap)
        {
            using var ms = new MemoryStream();
            bitmap.Save(ms, Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;

            var bi = new BitmapImage();
            bi.BeginInit();
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.StreamSource = ms;
            bi.EndInit();
            bi.Freeze();
            return bi;
        }

        private void EnsurePdfLoaded()
        {
            if (_currentPdfPath == null || _pdfViewerDoc == null || PdfImage.Source == null)
                throw new InvalidOperationException("Load a PDF first.");
        }

        // =========================
        // Selection drag
        // =========================
        private void ClearBox_Click(object sender, RoutedEventArgs e)
        {
            ClearSelectionOnly();
            ViewModel.RectStatus = "Selection cleared.";
        }

        private void ClearSelectionOnly()
        {
            _isDrawing = false;
            _selectionRectShape = null;
            SelectionLayer.Children.Clear();

            ViewModel.ScreenRectText = "ScreenRect: (none)";
            ViewModel.PdfRectText = "PdfRect: (none)";
            _lastSelectionPdfRect = null;
            _lastSelectionText = "";
        }

        private void Overlay_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_pdfViewerDoc == null || PdfImage.Source == null) return;

            SelectionLayer.CaptureMouse();
            _isDrawing = true;

            _start = e.GetPosition(SelectionLayer);

            _selectionRectShape = new System.Windows.Shapes.Rectangle
            {
                Stroke = Brushes.Magenta,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(50, 255, 0, 255))
            };

            Canvas.SetLeft(_selectionRectShape, _start.X);
            Canvas.SetTop(_selectionRectShape, _start.Y);
            _selectionRectShape.Width = 0;
            _selectionRectShape.Height = 0;

            SelectionLayer.Children.Clear();
            SelectionLayer.Children.Add(_selectionRectShape);

            ViewModel.RectStatus = _editMode == EditMode.RedrawRectForItem
                ? "Redraw mode: drag to replace selected mapping…"
                : "Drawing selection…";
        }

        private void Overlay_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDrawing || _selectionRectShape == null) return;

            var cur = e.GetPosition(SelectionLayer);

            double x = Math.Min(cur.X, _start.X);
            double y = Math.Min(cur.Y, _start.Y);
            double w = Math.Abs(cur.X - _start.X);
            double h = Math.Abs(cur.Y - _start.Y);

            Canvas.SetLeft(_selectionRectShape, x);
            Canvas.SetTop(_selectionRectShape, y);
            _selectionRectShape.Width = w;
            _selectionRectShape.Height = h;

            UpdateReadouts(x, y, w, h, isFinal: false);
        }

        private void Overlay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDrawing) return;

            _isDrawing = false;
            SelectionLayer.ReleaseMouseCapture();

            if (_selectionRectShape == null) return;

            double x = Canvas.GetLeft(_selectionRectShape);
            double y = Canvas.GetTop(_selectionRectShape);
            double w = _selectionRectShape.Width;
            double h = _selectionRectShape.Height;

            if (w < 5 || h < 5)
            {
                SelectionLayer.Children.Clear();
                ViewModel.RectStatus = "Selection too small.";
                _lastSelectionPdfRect = null;
                _lastSelectionText = "";
                return;
            }

            UpdateReadouts(x, y, w, h, isFinal: true);

            if (_editMode == EditMode.RedrawRectForItem && _pendingEditItemId != Guid.Empty && _lastSelectionPdfRect != null)
            {
                var item = FindItemById(_pendingEditItemId);
                if (item != null)
                {
                    var oldRect = item.Rect;
                    var newRect = _lastSelectionPdfRect.Value;

                    PushAndDo(new EditRectCommand(this, _pendingEditItemId, oldRect, newRect));
                    ExitRedrawMode();
                }
                else
                {
                    ExitRedrawMode();
                    ViewModel.RectStatus = "Redraw failed: item not found.";
                }
            }
            else
            {
                ViewModel.RectStatus = "Selection captured.";
            }
        }

        private void UpdateReadouts(double x, double y, double w, double h, bool isFinal)
        {
            ViewModel.ScreenRectText =
                $"ScreenRect(px): x={x:0.0}, y={y:0.0}, w={w:0.0}, h={h:0.0}";

            double imgW = SelectionLayer.Width;
            double imgH = SelectionLayer.Height;
            if (imgW <= 0 || imgH <= 0) return;

            double sx = _pageSizePoints.Width / imgW;
            double sy = _pageSizePoints.Height / imgH;

            double pdfX = x * sx;
            double pdfY = y * sy;
            double pdfW = w * sx;
            double pdfH = h * sy;

            ViewModel.PdfRectText =
                $"PdfRect(pt, top-left origin): x={pdfX:0.00}, y={pdfY:0.00}, w={pdfW:0.00}, h={pdfH:0.00}" +
                (isFinal ? "  [FINAL]" : "");

            if (isFinal)
                _lastSelectionPdfRect = new PdfRect(pdfX, pdfY, pdfW, pdfH);

            if (isFinal && _currentPdfPath != null)
            {
                try
                {
                    _lastSelectionText = ExtractTextInRegion(_currentPdfPath, _pageIndex, pdfX, pdfY, pdfW, pdfH);
                    ViewModel.NotesText = $"Extracted Text:\n{_lastSelectionText}";
                }
                catch (Exception ex)
                {
                    _lastSelectionText = "";
                    ViewModel.NotesText = $"Extraction error: {ex.Message}";
                }
            }
        }

        // =========================
        // Extraction (PdfPig)
        // =========================
        private string ExtractTextInRegion(string pdfPath, int pageIndex, double pdfX, double pdfY, double pdfW, double pdfH)
        {
            using var document = PigDoc.Open(pdfPath);
            var page = document.GetPage(pageIndex + 1);

            double pageHeight = page.Height;
            double bottomY = pageHeight - (pdfY + pdfH);

            double x0 = pdfX;
            double y0 = bottomY;
            double x1 = pdfX + pdfW;
            double y1 = bottomY + pdfH;

            var words = page.GetWords();

            var selected = words
                .Where(w =>
                {
                    double cx = (w.BoundingBox.Left + w.BoundingBox.Right) / 2.0;
                    double cy = (w.BoundingBox.Bottom + w.BoundingBox.Top) / 2.0;
                    return cx >= x0 && cx <= x1 && cy >= y0 && cy <= y1;
                })
                .OrderByDescending(w => w.BoundingBox.Top)
                .ThenBy(w => w.BoundingBox.Left)
                .Select(w => w.Text);

            return string.Join(" ", selected);
        }

        // =========================
        // Commands
        // =========================
        private sealed class AddItemCommand : IUndoableCommand
        {
            private readonly MainWindow _w;
            private readonly DesignerItem _item;
            public string Name => $"Add {_item.Kind}: {_item.Name}";
            public AddItemCommand(MainWindow w, DesignerItem item) { _w = w; _item = item; }
            public void Do() => _w.AddItemInternal(_item);
            public void Undo() => _w.RemoveItemInternal(_item.Id);
        }

        private sealed class DeleteItemCommand : IUndoableCommand
        {
            private readonly MainWindow _w;
            private readonly DesignerItem _snapshot;
            public string Name => $"Delete {_snapshot.Kind}: {_snapshot.Name}";
            public DeleteItemCommand(MainWindow w, DesignerItem snapshot) { _w = w; _snapshot = snapshot; }
            public void Do() => _w.RemoveItemInternal(_snapshot.Id);
            public void Undo() => _w.AddItemInternal(_snapshot);
        }

        private sealed class RenameCommand : IUndoableCommand
        {
            private readonly MainWindow _w;
            private readonly Guid _id;
            private readonly string _oldName;
            private readonly string _newName;
            public string Name => $"Rename to '{_newName}'";
            public RenameCommand(MainWindow w, Guid id, string oldName, string newName)
            { _w = w; _id = id; _oldName = oldName; _newName = newName; }
            public void Do() => _w.RenameItemInternal(_id, _newName);
            public void Undo() => _w.RenameItemInternal(_id, _oldName);
        }

        private sealed class EditRectCommand : IUndoableCommand
        {
            private readonly MainWindow _w;
            private readonly Guid _id;
            private readonly PdfRect _old;
            private readonly PdfRect _new;
            public string Name => "Edit region";
            public EditRectCommand(MainWindow w, Guid id, PdfRect oldRect, PdfRect newRect)
            { _w = w; _id = id; _old = oldRect; _new = newRect; }
            public void Do() => _w.EditRectInternal(_id, _new);
            public void Undo() => _w.EditRectInternal(_id, _old);
        }

        private sealed class ClearFormCommand : IUndoableCommand
        {
            private readonly MainWindow _w;
            private readonly string _snapshotJson;
            public string Name => "Clear Form";
            public ClearFormCommand(MainWindow w)
            { _w = w; _snapshotJson = w.ExportDesignerStateJson(); }
            public void Do() => _w.ClearFormInternal();
            public void Undo() => _w.ImportDesignerStateJson(_snapshotJson);
        }

        // =========================
        // Buttons
        // =========================
        private void ClearForm_Click(object sender, RoutedEventArgs e) => PushAndDo(new ClearFormCommand(this));

        private void AddFieldFromSelection_Click(object sender, RoutedEventArgs e)
        {
            try { EnsurePdfLoaded(); } catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            if (_lastSelectionPdfRect == null)
            {
                MessageBox.Show("Draw a selection box first.");
                return;
            }

            string name = PromptText("Field name:", "New Field");
            if (string.IsNullOrWhiteSpace(name)) return;

            var item = new FieldDef
            {
                Id = Guid.NewGuid(),
                Name = name.Trim(),
                PageIndex = _pageIndex,
                Rect = _lastSelectionPdfRect.Value,
                LastPreview = _lastSelectionText ?? ""
            };

            PushAndDo(new AddItemCommand(this, item));
        }

        private void AddTable_Click(object sender, RoutedEventArgs e)
        {
            try { EnsurePdfLoaded(); } catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            string name = PromptText("Table name:", "New Table");
            if (string.IsNullOrWhiteSpace(name)) return;

            var t = new TableDef
            {
                Id = Guid.NewGuid(),
                Name = name.Trim(),
                PageIndex = _pageIndex,
                Rect = new PdfRect(0, 0, 0, 0),
                SkipHeaderRows = 1,
                YTolerance = 2.5
            };

            PushAndDo(new AddItemCommand(this, t));
            SelectDesignerItem(t);
            ViewModel.NotesText = "Table added. Now draw table region → Set Table Region.";
        }

        private void SetTableRegionFromSelection_Click(object sender, RoutedEventArgs e)
        {
            try { EnsurePdfLoaded(); } catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            if (_lastSelectionPdfRect == null)
            {
                MessageBox.Show("Draw a selection box first.");
                return;
            }

            if (ViewModel.SelectedTable == null)
            {
                MessageBox.Show("Select a table first (left panel).");
                return;
            }

            var t = ViewModel.SelectedTable;
            PushAndDo(new EditRectCommand(this, t.Id, t.Rect, _lastSelectionPdfRect.Value));
            ViewModel.NotesText = "Table region set. Now draw a vertical band → Add Column.";
        }

        private void AddColumnFromSelection_Click(object sender, RoutedEventArgs e)
        {
            try { EnsurePdfLoaded(); } catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            if (_lastSelectionPdfRect == null)
            {
                MessageBox.Show("Draw a vertical band selection first.");
                return;
            }

            if (ViewModel.SelectedTable == null)
            {
                MessageBox.Show("Select a table first (left panel).");
                return;
            }

            var t = ViewModel.SelectedTable;
            if (t.Rect.W <= 1 || t.Rect.H <= 1)
            {
                MessageBox.Show("Set the table region first.");
                return;
            }

            string name = PromptText("Column name:", "New Column");
            if (string.IsNullOrWhiteSpace(name)) return;

            var sel = _lastSelectionPdfRect.Value;
            var table = t.Rect;

            double x0 = Math.Max(table.X, sel.X);
            double x1 = Math.Min(table.X + table.W, sel.X + sel.W);
            if (x1 <= x0 + 1)
            {
                MessageBox.Show("Column selection must overlap the table region.");
                return;
            }

            var col = new ColumnDef
            {
                Id = Guid.NewGuid(),
                Name = name.Trim(),
                PageIndex = t.PageIndex,
                Rect = new PdfRect(x0, table.Y, x1 - x0, table.H),
                ParentTableId = t.Id,
                Type = "string"
            };

            PushAndDo(new AddItemCommand(this, col));
            SelectDesignerItem(col);
        }

        private void EditSelected_Click(object sender, RoutedEventArgs e)
        {
            var sel = ViewModel.SelectedItem;
            if (sel == null)
            {
                MessageBox.Show("Select an item first (overlay or dropdown).");
                return;
            }

            var action = PromptChoice(
                title: $"Edit {sel.Kind}: {sel.Name}",
                message: "Choose an edit action:",
                choices: new[] { "Rename", "Redraw Region/Band", "Cancel" });

            if (action == "Rename")
            {
                string newName = PromptText("New name:", "Rename", sel.Name);
                if (string.IsNullOrWhiteSpace(newName)) return;
                newName = newName.Trim();
                if (newName == sel.Name) return;

                PushAndDo(new RenameCommand(this, sel.Id, sel.Name, newName));
            }
            else if (action == "Redraw Region/Band")
            {
                EnterRedrawMode(sel.Id);
            }
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var sel = ViewModel.SelectedItem;
            if (sel == null) { MessageBox.Show("Select an item first."); return; }

            if (MessageBox.Show($"Delete {sel.Kind}: '{sel.Name}'?", "Confirm Delete", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            PushAndDo(new DeleteItemCommand(this, CloneItem(sel)));
        }

        private void PreviewExport_Click(object sender, RoutedEventArgs e)
        {
            try { EnsurePdfLoaded(); } catch (Exception ex) { MessageBox.Show(ex.Message); return; }
            if (_currentPdfPath == null) return;

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("=== PREVIEW EXPORT ===");
                sb.AppendLine($"PDF: {IOPath.GetFileName(_currentPdfPath)}");
                sb.AppendLine();

                if (ViewModel.Fields.Count > 0)
                {
                    sb.AppendLine("[FIELDS]");
                    foreach (var f in ViewModel.Fields.OrderBy(x => x.Name))
                    {
                        string v = ExtractTextInRegion(_currentPdfPath, f.PageIndex, f.Rect.X, f.Rect.Y, f.Rect.W, f.Rect.H);
                        sb.AppendLine($"{f.Name}: {v}");
                    }
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine("[FIELDS] (none)");
                    sb.AppendLine();
                }

                if (ViewModel.Tables.Count > 0)
                {
                    foreach (var t in ViewModel.Tables.OrderBy(x => x.Name))
                    {
                        sb.AppendLine($"[TABLE] {t.Name}");
                        if (t.Rect.W <= 1 || t.Rect.H <= 1)
                        {
                            sb.AppendLine("  (table region not set)");
                            sb.AppendLine();
                            continue;
                        }

                        var cols = t.Columns.OrderBy(c => c.Rect.X).ToList();
                        if (cols.Count == 0)
                        {
                            sb.AppendLine("  (no columns)");
                            sb.AppendLine();
                            continue;
                        }

                        var rows = ExtractTableRows(_currentPdfPath, t.PageIndex, t.Rect, cols, t.SkipHeaderRows, t.YTolerance);

                        sb.AppendLine($"  Rows: {rows.Count} (showing first 10)");
                        sb.AppendLine("  " + string.Join(" | ", cols.Select(c => c.Name)));
                        sb.AppendLine("  " + new string('-', 90));

                        int take = Math.Min(10, rows.Count);
                        for (int i = 0; i < take; i++)
                        {
                            var row = rows[i];
                            sb.AppendLine("  " + string.Join(" | ", cols.Select(c => row.TryGetValue(c.Name, out var v) ? v : "")));
                        }

                        sb.AppendLine();
                    }
                }
                else
                {
                    sb.AppendLine("[TABLES] (none)");
                }

                ViewModel.NotesText = sb.ToString();
                ViewModel.RectStatus = "Preview complete.";
            }
            catch (Exception ex)
            {
                ViewModel.NotesText = $"Preview error: {ex.Message}";
                ViewModel.RectStatus = "Preview failed.";
            }
        }

        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Template JSON (*.json)|*.json",
                Title = "Save Template",
                FileName = "template.json"
            };
            if (dlg.ShowDialog() != true) return;

            var template = new TemplateDef
            {
                SchemaVersion = 2,
                Name = "Template",
                TextOnly = true,
                Fields = ViewModel.Fields.Select(f => (FieldDef)CloneItem(f)).ToList(),
                Tables = ViewModel.Tables.Select(t => t.ToSerializable()).ToList()
            };

            var opts = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };

            File.WriteAllText(dlg.FileName, JsonSerializer.Serialize(template, opts));
            ViewModel.RectStatus = $"Saved template: {IOPath.GetFileName(dlg.FileName)}";
        }

        private void LoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Template JSON (*.json)|*.json|All files (*.*)|*.*",
                Title = "Load Template JSON"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var json = File.ReadAllText(dlg.FileName);
                var template = JsonSerializer.Deserialize<TemplateDef>(json);

                if (template == null)
                {
                    MessageBox.Show("Template parse failed (null).");
                    return;
                }

                ViewModel.Fields.Clear();
                ViewModel.Tables.Clear();
                ViewModel.SelectedItem = null;
                ViewModel.SelectedField = null;
                ViewModel.SelectedTable = null;
                ViewModel.SelectedColumn = null;

                if (template.Fields != null)
                    foreach (var f in template.Fields) ViewModel.Fields.Add(f);

                if (template.Tables != null)
                    foreach (var t in template.Tables) ViewModel.Tables.Add(TableDef.FromSerializable(t));

                RebuildAllItems();
                RedrawPersistedOverlays();

                ViewModel.RectStatus = $"Loaded template: {IOPath.GetFileName(dlg.FileName)}";
                ViewModel.NotesText =
                    $"Template loaded.\nSchemaVersion: {template.SchemaVersion}\n" +
                    $"Fields: {ViewModel.Fields.Count}\nTables: {ViewModel.Tables.Count}\n\n" +
                    "Tip: Open a PDF that matches the template to visually verify overlays, then run Batch → output.csv.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Load template failed:\n{ex.Message}");
            }
        }

        private void BatchProcess_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.Fields.Count == 0 && ViewModel.Tables.Count == 0)
            {
                MessageBox.Show("No template loaded/defined. Create or load a template first.");
                return;
            }

            var selectedPdfs = PickPdfFiles("Select PDFs to batch (tip: open the folder, press Ctrl+A, then Open)");
            if (selectedPdfs == null || selectedPdfs.Count == 0) return;

            string? outputDir = PickOutputFolderViaCsvPath("Choose where to save output.csv");
            if (string.IsNullOrWhiteSpace(outputDir)) return;

            try
            {
                var pdfs = selectedPdfs;
if (pdfs.Count == 0)
                {
                    MessageBox.Show("No PDFs selected.");
                    return;
                }

                Directory.CreateDirectory(outputDir!);
var fieldNames = ViewModel.Fields
                    .Select(f => f.Name)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var tableColNames = new List<string>();
                foreach (var t in ViewModel.Tables.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
                {
                    foreach (var c in t.Columns.OrderBy(c => c.Rect.X))
                    {
                        var key = $"{t.Name}__{c.Name}";
                        if (!tableColNames.Contains(key, StringComparer.OrdinalIgnoreCase))
                            tableColNames.Add(key);
                    }
                }

                var header = new List<string> { "SourceFile", "TableName", "RowIndex" };
                header.AddRange(fieldNames);
                header.AddRange(tableColNames);

                string outCsv = IOPath.Combine(outputDir, "output.csv");
                using var sw = new StreamWriter(outCsv, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                WriteCsvRow(sw, header.ToArray());

                int totalRowsWritten = 0;

                foreach (var pdf in pdfs)
                {
                    var fieldValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var fname in fieldNames)
                    {
                        var f = ViewModel.Fields.FirstOrDefault(x => string.Equals(x.Name, fname, StringComparison.OrdinalIgnoreCase));
                        if (f == null) { fieldValues[fname] = ""; continue; }

                        try
                        {
                            fieldValues[fname] = ExtractTextInRegion(pdf, f.PageIndex, f.Rect.X, f.Rect.Y, f.Rect.W, f.Rect.H);
                        }
                        catch { fieldValues[fname] = ""; }
                    }

                    bool wroteAnyTableRows = false;

                    foreach (var t in ViewModel.Tables.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        var cols = t.Columns.OrderBy(c => c.Rect.X).ToList();
                        if (cols.Count == 0) continue;
                        if (t.Rect.W <= 1 || t.Rect.H <= 1) continue;

                        List<Dictionary<string, string>> rows;
                        try
                        {
                            rows = ExtractTableRows(pdf, t.PageIndex, t.Rect, cols, t.SkipHeaderRows, t.YTolerance);
                        }
                        catch
                        {
                            rows = new List<Dictionary<string, string>>();
                        }

                        for (int i = 0; i < rows.Count; i++)
                        {
                            wroteAnyTableRows = true;
                            var r = rows[i];

                            var line = new List<string>
                            {
                                IOPath.GetFileName(pdf),
                                t.Name,
                                (i + 1).ToString()
                            };

                            foreach (var fname in fieldNames)
                                line.Add(fieldValues.TryGetValue(fname, out var fv) ? fv : "");

                            var colMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            foreach (var c in cols)
                            {
                                var key = $"{t.Name}__{c.Name}";
                                colMap[key] = r.TryGetValue(c.Name, out var v) ? v : "";
                            }

                            foreach (var k in tableColNames)
                                line.Add(colMap.TryGetValue(k, out var v) ? v : "");

                            WriteCsvRow(sw, line.ToArray());
                            totalRowsWritten++;
                        }
                    }

                    if (!wroteAnyTableRows)
                    {
                        var line = new List<string>
                        {
                            IOPath.GetFileName(pdf),
                            "",
                            ""
                        };

                        foreach (var fname in fieldNames)
                            line.Add(fieldValues.TryGetValue(fname, out var fv) ? fv : "");

                        foreach (var _ in tableColNames)
                            line.Add("");

                        WriteCsvRow(sw, line.ToArray());
                        totalRowsWritten++;
                    }
                }

                ViewModel.RectStatus = "Batch complete.";
                ViewModel.NotesText =
                    "Batch export complete (single combined CSV).\n\n" +
                    $"Input:  {IOPath.GetDirectoryName(pdfs[0])} (selected {pdfs.Count} PDFs)\n" +
                    $"Output: {outputDir!}\n\n" +
                    $"File written:\n- {IOPath.GetFileName(outCsv)}\n\n" +
                    $"Rows written: {totalRowsWritten}\n\n" +
                    "Notes:\n" +
                    "- Table columns are prefixed as <TableName>__<ColumnName> to avoid collisions.\n" +
                    "- Field values are repeated on every emitted row for that PDF.\n";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Batch failed:\n{ex.Message}");
            }
        }

        private static void WriteCsvRow(StreamWriter sw, string[] fields)
        {
            string Escape(string s)
            {
                s ??= "";
                bool needsQuotes = s.Contains(',') || s.Contains('"') || s.Contains('\n') || s.Contains('\r');
                if (s.Contains('"')) s = s.Replace("\"", "\"\"");
                return needsQuotes ? $"\"{s}\"" : s;
            }

            sw.WriteLine(string.Join(",", fields.Select(Escape)));
        }

        private static string SafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            return string.IsNullOrWhiteSpace(cleaned) ? "table" : cleaned.Trim();
        }

        private static List<string>? PickPdfFiles(string title)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = title,
                Filter = "PDF files (*.pdf)|*.pdf",
                Multiselect = true
            };
            if (dlg.ShowDialog() != true) return null;

            return dlg.FileNames
                      .Where(f => !string.IsNullOrWhiteSpace(f) && File.Exists(f))
                      .Distinct(StringComparer.OrdinalIgnoreCase)
                      .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                      .ToList();
        }

        private static string? PickOutputFolderViaCsvPath(string title)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = title,
                Filter = "CSV (*.csv)|*.csv",
                FileName = "output.csv"
            };
            if (dlg.ShowDialog() != true) return null;
            return IOPath.GetDirectoryName(dlg.FileName);
        }

        
        
        
        // =========================
        // Table extraction
        // =========================
        private sealed class WordSpan
        {
            public string Text = "";
            public double L, R, B, T;
        }

        private sealed class RowGroup
        {
            public double CenterY;
            public double AvgTop;
            public List<WordSpan> Words { get; } = new();
            public void Recompute()
            {
                if (Words.Count == 0) { AvgTop = CenterY; return; }
                AvgTop = Words.Average(w => w.T);
                CenterY = Words.Average(w => (w.B + w.T) / 2.0);
            }
        }

        private List<Dictionary<string, string>> ExtractTableRows(
            string pdfPath,
            int pageIndex,
            PdfRect tableRectTopLeft,
            List<ColumnDef> columns,
            int skipHeaderRows,
            double yTolerance)
        {
            using var document = PigDoc.Open(pdfPath);
            var page = document.GetPage(pageIndex + 1);

            double pageHeight = page.Height;
            double tableBottomY = pageHeight - (tableRectTopLeft.Y + tableRectTopLeft.H);

            double tx0 = tableRectTopLeft.X;
            double ty0 = tableBottomY;
            double tx1 = tableRectTopLeft.X + tableRectTopLeft.W;
            double ty1 = tableBottomY + tableRectTopLeft.H;

            var words = page.GetWords()
                .Select(w => new WordSpan
                {
                    Text = w.Text,
                    L = w.BoundingBox.Left,
                    R = w.BoundingBox.Right,
                    B = w.BoundingBox.Bottom,
                    T = w.BoundingBox.Top
                })
                .Where(ws =>
                {
                    double cx = (ws.L + ws.R) / 2.0;
                    double cy = (ws.B + ws.T) / 2.0;
                    return cx >= tx0 && cx <= tx1 && cy >= ty0 && cy <= ty1;
                })
                .ToList();

            if (words.Count == 0) return new List<Dictionary<string, string>>();

            var rowGroups = ClusterRowsByY(words, yTolerance)
                .OrderByDescending(g => g.AvgTop)
                .ToList();

            if (skipHeaderRows > 0 && rowGroups.Count > skipHeaderRows)
                rowGroups = rowGroups.Skip(skipHeaderRows).ToList();

            var cols = columns.OrderBy(c => c.Rect.X).ToList();
            var output = new List<Dictionary<string, string>>();

            foreach (var rg in rowGroups)
            {
                var rowDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var col in cols)
                {
                    double cx0 = col.Rect.X;
                    double cx1 = col.Rect.X + col.Rect.W;

                    var inCol = rg.Words
                        .Where(w =>
                        {
                            double cwx = (w.L + w.R) / 2.0;
                            return cwx >= cx0 && cwx <= cx1;
                        })
                        .OrderBy(w => w.L)
                        .Select(w => w.Text);

                    rowDict[col.Name] = string.Join(" ", inCol).Trim();
                }

                if (rowDict.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                    output.Add(rowDict);
            }

            return output;
        }

        private List<RowGroup> ClusterRowsByY(List<WordSpan> words, double yTolerance)
        {
            var sorted = words
                .OrderByDescending(w => (w.B + w.T) / 2.0)
                .ToList();

            var groups = new List<RowGroup>();

            foreach (var w in sorted)
            {
                double cy = (w.B + w.T) / 2.0;
                RowGroup? found = null;

                foreach (var g in groups)
                {
                    if (Math.Abs(cy - g.CenterY) <= yTolerance)
                    {
                        found = g;
                        break;
                    }
                }

                if (found == null)
                {
                    var g = new RowGroup { CenterY = cy };
                    g.Words.Add(w);
                    g.Recompute();
                    groups.Add(g);
                }
                else
                {
                    found.Words.Add(w);
                    found.Recompute();
                }
            }

            return groups;
        }

        // =========================
        // Internal state ops
        // =========================
        private void AddItemInternal(DesignerItem item)
        {
            switch (item.Kind)
            {
                case ItemKind.Field:
                    ViewModel.Fields.Add((FieldDef)item);
                    break;

                case ItemKind.Table:
                    ViewModel.Tables.Add((TableDef)item);
                    break;

                case ItemKind.Column:
                    {
                        var c = (ColumnDef)item;
                        var table = ViewModel.Tables.FirstOrDefault(t => t.Id == c.ParentTableId);
                        if (table == null) throw new InvalidOperationException("Parent table not found.");

                        table.Columns.Add(c);

                        var sorted = table.Columns.OrderBy(x => x.Rect.X).ToList();
                        table.Columns.Clear();
                        foreach (var s in sorted) table.Columns.Add(s);

                        break;
                    }
            }

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        private void RemoveItemInternal(Guid id)
        {
            var f = ViewModel.Fields.FirstOrDefault(x => x.Id == id);
            if (f != null)
            {
                ViewModel.Fields.Remove(f);
                if (ViewModel.SelectedItem?.Id == id) ViewModel.SelectedItem = null;
                RebuildAllItems();
                RedrawPersistedOverlays();
                return;
            }

            var t = ViewModel.Tables.FirstOrDefault(x => x.Id == id);
            if (t != null)
            {
                ViewModel.Tables.Remove(t);
                if (ViewModel.SelectedItem?.Id == id) ViewModel.SelectedItem = null;
                RebuildAllItems();
                RedrawPersistedOverlays();
                return;
            }

            foreach (var table in ViewModel.Tables)
            {
                var c = table.Columns.FirstOrDefault(x => x.Id == id);
                if (c != null)
                {
                    table.Columns.Remove(c);
                    if (ViewModel.SelectedItem?.Id == id) ViewModel.SelectedItem = null;
                    RebuildAllItems();
                    RedrawPersistedOverlays();
                    return;
                }
            }

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        private void RenameItemInternal(Guid id, string newName)
        {
            var item = FindItemById(id);
            if (item == null) return;

            item.Name = newName;
            item.RaiseChanged();

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        private void EditRectInternal(Guid id, PdfRect rect)
        {
            var item = FindItemById(id);
            if (item == null) return;

            item.Rect = rect;
            item.RaiseChanged();

            if (item is TableDef t && rect.W > 1 && rect.H > 1)
            {
                foreach (var c in t.Columns)
                {
                    c.Rect = new PdfRect(c.Rect.X, rect.Y, c.Rect.W, rect.H);
                    c.RaiseChanged();
                }
            }

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        private void ClearFormInternal()
        {
            ViewModel.Fields.Clear();
            ViewModel.Tables.Clear();
            ViewModel.SelectedItem = null;
            ViewModel.SelectedField = null;
            ViewModel.SelectedTable = null;
            ViewModel.SelectedColumn = null;

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        // =========================
        // Designer state snapshot
        // =========================
        private string ExportDesignerStateJson()
        {
            var state = new DesignerState
            {
                Fields = ViewModel.Fields.Select(f => (FieldDef)CloneItem(f)).ToList(),
                Tables = ViewModel.Tables.Select(t => t.ToSerializable()).ToList()
            };

            var opts = new JsonSerializerOptions
            {
                WriteIndented = false,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };

            return JsonSerializer.Serialize(state, opts);
        }

        private void ImportDesignerStateJson(string json)
        {
            var state = JsonSerializer.Deserialize<DesignerState>(json);
            if (state == null) return;

            ViewModel.Fields.Clear();
            ViewModel.Tables.Clear();

            foreach (var f in state.Fields)
                ViewModel.Fields.Add(f);

            foreach (var t in state.Tables)
                ViewModel.Tables.Add(TableDef.FromSerializable(t));

            RebuildAllItems();
            RedrawPersistedOverlays();
        }

        // =========================
        // Selection / item picking
        // =========================
        private DesignerItem? FindItemById(Guid id)
        {
            var f = ViewModel.Fields.FirstOrDefault(x => x.Id == id);
            if (f != null) return f;

            var t = ViewModel.Tables.FirstOrDefault(x => x.Id == id);
            if (t != null) return t;

            foreach (var table in ViewModel.Tables)
            {
                var c = table.Columns.FirstOrDefault(x => x.Id == id);
                if (c != null) return c;
            }

            return null;
        }

        private void SelectDesignerItem(DesignerItem? item)
        {
            if (item == null) return;

            if (ReferenceEquals(ViewModel.SelectedItem, item))
            {
                ViewModel.SelectedItemSummary = item.GetSummary();
                ViewModel.SelectedTableSummary = ViewModel.SelectedTable?.GetSummary() ?? "Select a table to see details.";
                HighlightSelectedOverlay(item.Id);
                return;
            }

            if (item is FieldDef f)
            {
                ViewModel.SelectedField = f;
                ViewModel.SelectedTable = null;
                ViewModel.SelectedColumn = null;
            }
            else if (item is TableDef t)
            {
                ViewModel.SelectedTable = t;
                ViewModel.SelectedField = null;
                ViewModel.SelectedColumn = null;
            }
            else if (item is ColumnDef c)
            {
                var parent = ViewModel.Tables.FirstOrDefault(x => x.Id == c.ParentTableId);
                ViewModel.SelectedTable = parent;
                ViewModel.SelectedColumn = c;
                ViewModel.SelectedField = null;
            }

            ViewModel.SelectedItem = item;
            ViewModel.SelectedItemSummary = item.GetSummary();
            ViewModel.SelectedTableSummary = ViewModel.SelectedTable?.GetSummary() ?? "Select a table to see details.";

            HighlightSelectedOverlay(item.Id);
        }

        private void HighlightSelectedOverlay(Guid id)
        {
            foreach (var kv in _visuals)
            {
                var vis = kv.Value;
                if (kv.Key == id)
                {
                    vis.Rect.StrokeThickness = 3.5;
                    vis.Label.FontWeight = FontWeights.Bold;
                }
                else
                {
                    vis.Rect.StrokeThickness = 2.0;
                    vis.Label.FontWeight = FontWeights.Normal;
                }
            }
        }

        // =========================
        // Overlay drawing (persisted)
        // =========================
        private void RedrawPersistedOverlays()
        {
            PersistedLayer.Children.Clear();
            _visuals.Clear();

            if (PdfImage.Source == null) return;

            foreach (var f in ViewModel.Fields)
                AddOverlayVisualForItem(f);

            foreach (var t in ViewModel.Tables)
            {
                AddOverlayVisualForItem(t);
                foreach (var c in t.Columns)
                    AddOverlayVisualForItem(c);
            }

            if (ViewModel.SelectedItem != null)
                HighlightSelectedOverlay(ViewModel.SelectedItem.Id);
        }

        private void AddOverlayVisualForItem(DesignerItem item)
        {
            if (item.Rect.W <= 1 || item.Rect.H <= 1) return;

            var (stroke, fill) = GetBrushes(item.Kind);

            var px = PdfToScreenRect(item.Rect);

            var rect = new System.Windows.Shapes.Rectangle
            {
                Stroke = stroke,
                Fill = fill,
                StrokeThickness = 2,
                RadiusX = 2,
                RadiusY = 2,
                IsHitTestVisible = true,
                Tag = item.Id
            };

            Canvas.SetLeft(rect, px.X);
            Canvas.SetTop(rect, px.Y);
            rect.Width = px.W;
            rect.Height = px.H;

            rect.MouseLeftButtonDown += PersistedOverlay_MouseDown;

            var label = new TextBlock
            {
                Text = item.Kind == ItemKind.Column ? $"Col: {item.Name}"
                     : item.Kind == ItemKind.Table ? $"Table: {item.Name}"
                     : $"Field: {item.Name}",
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromArgb(130, 0, 0, 0)),
                Padding = new Thickness(4, 2, 4, 2),
                FontSize = 12,
                Tag = item.Id,
                IsHitTestVisible = true
            };

            Canvas.SetLeft(label, px.X);
            Canvas.SetTop(label, Math.Max(0, px.Y - 20));
            label.MouseLeftButtonDown += PersistedOverlay_MouseDown;

            PersistedLayer.Children.Add(rect);
            PersistedLayer.Children.Add(label);

            _visuals[item.Id] = new OverlayVisual { Id = item.Id, Rect = rect, Label = label, Kind = item.Kind };
        }

        private void PersistedOverlay_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is Guid id)
            {
                var item = FindItemById(id);
                if (item != null)
                {
                    SelectDesignerItem(item);
                    e.Handled = true;
                }
            }
        }

        private (Brush stroke, Brush fill) GetBrushes(ItemKind kind)
        {
            return kind switch
            {
                ItemKind.Field => (Brushes.DeepSkyBlue, new SolidColorBrush(Color.FromArgb(45, 30, 144, 255))),
                ItemKind.Table => (Brushes.LimeGreen, new SolidColorBrush(Color.FromArgb(45, 50, 205, 50))),
                ItemKind.Column => (Brushes.Orange, new SolidColorBrush(Color.FromArgb(35, 255, 165, 0))),
                _ => (Brushes.White, new SolidColorBrush(Color.FromArgb(35, 255, 255, 255))),
            };
        }

        private PdfRect PdfToScreenRect(PdfRect r)
        {
            double imgW = SelectionLayer.Width;
            double imgH = SelectionLayer.Height;
            if (imgW <= 0 || imgH <= 0) return r;

            double sx = imgW / _pageSizePoints.Width;
            double sy = imgH / _pageSizePoints.Height;

            return new PdfRect(
                X: r.X * sx,
                Y: r.Y * sy,
                W: r.W * sx,
                H: r.H * sy
            );
        }

        // =========================
        // All-items dropdown
        // =========================
        private void RebuildAllItems()
        {
            ViewModel.AllItems.Clear();

            foreach (var f in ViewModel.Fields.OrderBy(x => x.Name))
                ViewModel.AllItems.Add(f);

            foreach (var t in ViewModel.Tables.OrderBy(x => x.Name))
            {
                ViewModel.AllItems.Add(t);
                foreach (var c in t.Columns.OrderBy(x => x.Name))
                    ViewModel.AllItems.Add(c);
            }

            ViewModel.SelectedTableSummary = ViewModel.SelectedTable?.GetSummary() ?? "Select a table to see details.";
            if (ViewModel.SelectedItem != null)
                ViewModel.SelectedItemSummary = ViewModel.SelectedItem.GetSummary();
        }

        // =========================
        // Redraw mode helpers
        // =========================
        private void EnterRedrawMode(Guid id)
        {
            _editMode = EditMode.RedrawRectForItem;
            _pendingEditItemId = id;
            ViewModel.RectStatus = "Redraw mode: drag new rectangle to replace mapping.";
        }

        private void ExitRedrawMode()
        {
            _editMode = EditMode.Normal;
            _pendingEditItemId = Guid.Empty;
            ViewModel.RectStatus = "Ready.";
        }

        // =========================
        // Prompt helpers
        // =========================
        private static string PromptText(string label, string title, string initial = "")
        {
            var win = new Window
            {
                Title = title,
                Width = 460,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize
            };

            var panel = new DockPanel { Margin = new Thickness(10) };

            var tbLabel = new TextBlock { Text = label, Margin = new Thickness(0, 0, 0, 6) };
            DockPanel.SetDock(tbLabel, Dock.Top);
            panel.Children.Add(tbLabel);

            var tb = new TextBox { Margin = new Thickness(0, 0, 0, 10), Text = initial };
            DockPanel.SetDock(tb, Dock.Top);
            panel.Children.Add(tb);

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "OK", Width = 90, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = "Cancel", Width = 90, IsCancel = true };

            ok.Click += (_, __) => win.DialogResult = true;

            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            DockPanel.SetDock(buttons, Dock.Bottom);
            panel.Children.Add(buttons);

            win.Content = panel;

            if (Application.Current?.MainWindow != null)
                win.Owner = Application.Current.MainWindow;

            win.Loaded += (_, __) => { tb.Focus(); tb.SelectAll(); };

            return win.ShowDialog() == true ? tb.Text : "";
        }

        private static string PromptChoice(string title, string message, string[] choices)
        {
            var win = new Window
            {
                Title = title,
                Width = 420,
                Height = 220,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize
            };

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 10) });

            string picked = "Cancel";

            foreach (var c in choices)
            {
                var b = new Button { Content = c, Margin = new Thickness(0, 0, 0, 8), Height = 30 };
                b.Click += (_, __) =>
                {
                    picked = c;
                    win.DialogResult = true;
                };
                root.Children.Add(b);
            }

            win.Content = root;

            if (Application.Current?.MainWindow != null)
                win.Owner = Application.Current.MainWindow;

            win.ShowDialog();
            return picked;
        }

        // =========================
        // Cloning
        // =========================
        private static DesignerItem CloneItem(DesignerItem item)
        {
            return item.Kind switch
            {
                ItemKind.Field => new FieldDef
                {
                    Id = item.Id,
                    Name = item.Name,
                    PageIndex = item.PageIndex,
                    Rect = item.Rect,
                    LastPreview = (item as FieldDef)?.LastPreview
                },
                ItemKind.Table => new TableDef
                {
                    Id = item.Id,
                    Name = item.Name,
                    PageIndex = item.PageIndex,
                    Rect = item.Rect,
                    SkipHeaderRows = (item as TableDef)?.SkipHeaderRows ?? 1,
                    YTolerance = (item as TableDef)?.YTolerance ?? 2.5,
                    Columns = new ObservableCollection<ColumnDef>(
                        (item as TableDef)?.Columns.Select(c => (ColumnDef)CloneItem(c)).Cast<ColumnDef>().ToList() ?? new List<ColumnDef>())
                },
                ItemKind.Column => new ColumnDef
                {
                    Id = item.Id,
                    Name = item.Name,
                    PageIndex = item.PageIndex,
                    Rect = item.Rect,
                    ParentTableId = (item as ColumnDef)?.ParentTableId ?? Guid.Empty,
                    Type = (item as ColumnDef)?.Type ?? "string"
                },
                _ => throw new NotSupportedException()
            };
        }

        // =========================
        // Models
        // =========================
        public enum ItemKind { Field, Table, Column }

        public readonly record struct PdfRect(double X, double Y, double W, double H);

        public abstract class DesignerItem : INotifyPropertyChanged
        {
            public Guid Id { get; set; }
            public abstract ItemKind Kind { get; }
            public string Name { get; set; } = "";
            public int PageIndex { get; set; }
            public PdfRect Rect { get; set; }

            [JsonIgnore]
            public string DisplayName =>
                Kind == ItemKind.Column ? $"Column: {Name}" :
                Kind == ItemKind.Table ? $"Table: {Name}" :
                $"Field: {Name}";

            public virtual string GetSummary()
                => $"{DisplayName}\nPage {PageIndex + 1}\nRect: x={Rect.X:0.0} y={Rect.Y:0.0} w={Rect.W:0.0} h={Rect.H:0.0}";

            public event PropertyChangedEventHandler? PropertyChanged;
            public void RaiseChanged()
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Rect)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
            }
        }

        public class FieldDef : DesignerItem
        {
            public override ItemKind Kind => ItemKind.Field;
            public string? LastPreview { get; set; }

            [JsonIgnore]
            public string Summary => $"p{PageIndex + 1}  x={Rect.X:0.0} y={Rect.Y:0.0} w={Rect.W:0.0} h={Rect.H:0.0}";
        }

        public class ColumnDef : DesignerItem
        {
            public override ItemKind Kind => ItemKind.Column;
            public Guid ParentTableId { get; set; }
            public string Type { get; set; } = "string";

            [JsonIgnore]
            public string Summary => $"x0={Rect.X:0.0}  x1={(Rect.X + Rect.W):0.0}";
        }

        public class TableDef : DesignerItem
        {
            public override ItemKind Kind => ItemKind.Table;

            public int SkipHeaderRows { get; set; } = 1;
            public double YTolerance { get; set; } = 2.5;

            public ObservableCollection<ColumnDef> Columns { get; set; } = new();

            public override string GetSummary()
            {
                var baseSum = base.GetSummary();
                return baseSum + $"\nColumns: {Columns.Count}\nSkipHeaderRows: {SkipHeaderRows}  YTol: {YTolerance:0.##}";
            }

            public SerializableTable ToSerializable()
            {
                return new SerializableTable
                {
                    Id = Id,
                    Name = Name,
                    PageIndex = PageIndex,
                    Rect = Rect,
                    SkipHeaderRows = SkipHeaderRows,
                    YTolerance = YTolerance,
                    Columns = Columns.Select(c => new SerializableColumn
                    {
                        Id = c.Id,
                        Name = c.Name,
                        PageIndex = c.PageIndex,
                        Rect = c.Rect,
                        ParentTableId = c.ParentTableId,
                        Type = c.Type
                    }).ToList()
                };
            }

            public static TableDef FromSerializable(SerializableTable t)
            {
                var tvm = new TableDef
                {
                    Id = t.Id,
                    Name = t.Name,
                    PageIndex = t.PageIndex,
                    Rect = t.Rect,
                    SkipHeaderRows = t.SkipHeaderRows,
                    YTolerance = t.YTolerance,
                    Columns = new ObservableCollection<ColumnDef>()
                };

                foreach (var c in t.Columns)
                {
                    tvm.Columns.Add(new ColumnDef
                    {
                        Id = c.Id,
                        Name = c.Name,
                        PageIndex = c.PageIndex,
                        Rect = c.Rect,
                        ParentTableId = c.ParentTableId,
                        Type = c.Type
                    });
                }

                var sorted = tvm.Columns.OrderBy(x => x.Rect.X).ToList();
                tvm.Columns.Clear();
                foreach (var s in sorted) tvm.Columns.Add(s);

                return tvm;
            }
        }

        public class SerializableColumn
        {
            public Guid Id { get; set; }
            public string Name { get; set; } = "";
            public int PageIndex { get; set; }
            public PdfRect Rect { get; set; }
            public Guid ParentTableId { get; set; }
            public string Type { get; set; } = "string";
        }

        public class SerializableTable
        {
            public Guid Id { get; set; }
            public string Name { get; set; } = "";
            public int PageIndex { get; set; }
            public PdfRect Rect { get; set; }
            public int SkipHeaderRows { get; set; } = 1;
            public double YTolerance { get; set; } = 2.5;
            public List<SerializableColumn> Columns { get; set; } = new();
        }

        public class TemplateDef
        {
            public int SchemaVersion { get; set; } = 2;
            public string Name { get; set; } = "Template";
            public bool TextOnly { get; set; } = true;
            public List<FieldDef> Fields { get; set; } = new();
            public List<SerializableTable> Tables { get; set; } = new();
        }

        public class DesignerState
        {
            public List<FieldDef> Fields { get; set; } = new();
            public List<SerializableTable> Tables { get; set; } = new();
        }

        // =========================
        // VM
        // =========================
        public class Vm : INotifyPropertyChanged
        {
            private string _rectStatus = "";
            private string _screenRectText = "";
            private string _pdfRectText = "";
            private string _notesText = "";

            private FieldDef? _selectedField;
            private TableDef? _selectedTable;
            private ColumnDef? _selectedColumn;
            private DesignerItem? _selectedItem;

            private string _selectedItemSummary = "";
            private string _selectedTableSummary = "Select a table to see details.";

            public ObservableCollection<FieldDef> Fields { get; } = new();
            public ObservableCollection<TableDef> Tables { get; } = new();

            public ObservableCollection<DesignerItem> AllItems { get; } = new();

            public FieldDef? SelectedField { get => _selectedField; set { if (!ReferenceEquals(_selectedField, value)) { _selectedField = value; OnPropertyChanged(); } } }
            public TableDef? SelectedTable { get => _selectedTable; set { if (!ReferenceEquals(_selectedTable, value)) { _selectedTable = value; OnPropertyChanged(); OnPropertyChanged(nameof(SelectedTableSummary)); } } }
            public ColumnDef? SelectedColumn { get => _selectedColumn; set { if (!ReferenceEquals(_selectedColumn, value)) { _selectedColumn = value; OnPropertyChanged(); } } }

            public DesignerItem? SelectedItem { get => _selectedItem; set { if (!ReferenceEquals(_selectedItem, value)) { _selectedItem = value; OnPropertyChanged(); } } }

            public string SelectedItemSummary { get => _selectedItemSummary; set { if (_selectedItemSummary != value) { _selectedItemSummary = value; OnPropertyChanged(); } } }
            public string SelectedTableSummary { get => _selectedTableSummary; set { if (_selectedTableSummary != value) { _selectedTableSummary = value; OnPropertyChanged(); } } }

            public string RectStatus { get => _rectStatus; set { if (_rectStatus != value) { _rectStatus = value; OnPropertyChanged(); } } }
            public string ScreenRectText { get => _screenRectText; set { if (_screenRectText != value) { _screenRectText = value; OnPropertyChanged(); } } }
            public string PdfRectText { get => _pdfRectText; set { if (_pdfRectText != value) { _pdfRectText = value; OnPropertyChanged(); } } }
            public string NotesText { get => _notesText; set { if (_notesText != value) { _notesText = value; OnPropertyChanged(); } } }

            public event PropertyChangedEventHandler? PropertyChanged;
            private void OnPropertyChanged([CallerMemberName] string? name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
