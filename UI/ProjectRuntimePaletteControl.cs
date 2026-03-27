using System.Drawing;
using System.Windows.Forms;
using Plant3D.ProjectRuntimePalettes.Models;
using Plant3D.ProjectRuntimePalettes.Services;
using Plant3D.ProjectRuntimePalettes.Utilities;

namespace Plant3D.ProjectRuntimePalettes.UI;

public sealed class ProjectRuntimePaletteControl : UserControl
{
    private readonly SymbolPreviewService _previewService;
    private readonly ToolExecutionService _executionService;
    private readonly Label _headerLabel;
    private readonly Panel _toolbarPanel;
    private readonly Button _toggleTreeExpansionButton;
    private readonly Button _toggleTreePanelButton;
    private readonly TextBox _searchTextBox;
    private readonly ToolTip _toolTip;
    private readonly Label _statusLabel;
    private readonly SplitContainer _splitContainer;
    private readonly BufferedTreeView _classTree;
    private readonly BufferedListView _itemList;
    private readonly ImageList _imageList;

    private ProjectPaletteModel? _model;
    private bool _treeExpanded;
    private bool _treeExpandedBeforeSearch;
    private bool _searchActive;
    private bool _treePanelVisible = true;
    private bool _suppressSearchTextChanged;
    private int _savedSplitterDistance = 220;
    private string? _selectedNodeKey;

    public ProjectRuntimePaletteControl(SymbolPreviewService previewService, ToolExecutionService executionService)
    {
        _previewService = previewService;
        _executionService = executionService;

        BackColor = SystemColors.Control;
        Dock = DockStyle.Fill;

        _headerLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 46,
            Padding = new Padding(8, 8, 8, 4),
            Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true
        };

        _toolTip = new ToolTip();

        _toolbarPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 38,
            Padding = new Padding(6, 3, 6, 3)
        };
        _toolbarPanel.Resize += ToolbarPanel_Resize;

        _toggleTreeExpansionButton = CreateToolbarButton("+");
        _toggleTreeExpansionButton.Click += ToggleTreeExpansionButton_Click;

        _toggleTreePanelButton = CreateToolbarButton("<");
        _toggleTreePanelButton.Click += ToggleTreePanelButton_Click;

        _searchTextBox = new TextBox
        {
            Top = 1,
            Height = 28,
            PlaceholderText = "Search tree...",
            Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;
        _toolTip.SetToolTip(_searchTextBox, "Filter the class tree and keep all matching parent nodes visible.");

        _toolbarPanel.Controls.Add(_toggleTreeExpansionButton);
        _toolbarPanel.Controls.Add(_toggleTreePanelButton);
        _toolbarPanel.Controls.Add(_searchTextBox);
        LayoutToolbar();

        _statusLabel = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 22,
            Padding = new Padding(8, 0, 8, 0),
            TextAlign = ContentAlignment.MiddleLeft,
            BorderStyle = BorderStyle.FixedSingle
        };

        _splitContainer = new SplitContainer
        {
            Dock = DockStyle.Fill,
            FixedPanel = FixedPanel.Panel1,
            IsSplitterFixed = false,
            Panel1MinSize = 120,
            SplitterDistance = _savedSplitterDistance,
            BackColor = SystemColors.Control
        };

        _classTree = new BufferedTreeView
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            HideSelection = false,
            FullRowSelect = true,
            ShowLines = true,
            ShowRootLines = true,
            ShowPlusMinus = true,
            BackColor = SystemColors.Window,
            ForeColor = SystemColors.ControlText
        };
        _classTree.AfterSelect += ClassTree_AfterSelect;

        _imageList = new ImageList
        {
            ColorDepth = ColorDepth.Depth32Bit,
            ImageSize = new Size(SymbolPreviewService.PreviewPixelSize, SymbolPreviewService.PreviewPixelSize)
        };

        _itemList = new BufferedListView
        {
            Dock = DockStyle.Fill,
            View = View.LargeIcon,
            LargeImageList = _imageList,
            MultiSelect = false,
            HideSelection = false,
            HeaderStyle = ColumnHeaderStyle.None,
            ShowItemToolTips = true,
            BorderStyle = BorderStyle.FixedSingle,
            UseCompatibleStateImageBehavior = false,
            BackColor = SystemColors.Window,
            ForeColor = SystemColors.ControlText
        };
        _itemList.MouseUp += ItemList_MouseUp;

        _splitContainer.Panel1.Padding = new Padding(6);
        _splitContainer.Panel1.Controls.Add(_classTree);
        _splitContainer.Panel2.Padding = new Padding(6);
        _splitContainer.Panel2.Controls.Add(_itemList);

        Controls.Add(_splitContainer);
        Controls.Add(_statusLabel);
        Controls.Add(_toolbarPanel);
        Controls.Add(_headerLabel);

        UpdateToolbarButtons();
    }

    public void LoadModel(ProjectPaletteModel model)
    {
        _model = model;
        _headerLabel.Text = $"{model.Context.ProjectName}\r\n{model.ModeText}";

        _selectedNodeKey = null;
        _treeExpanded = false;
        _treeExpandedBeforeSearch = false;
        _searchActive = false;
        _treePanelVisible = true;
        _splitContainer.Panel1Collapsed = false;
        TryRestoreSplitterDistance();

        _suppressSearchTextChanged = true;
        _searchTextBox.Clear();
        _suppressSearchTextChanged = false;

        ApplyTreeFilter(preserveSelection: false);
        UpdateToolbarButtons();
    }

    private static Button CreateToolbarButton(string text)
    {
        return new Button
        {
            Width = 30,
            Height = 28,
            Left = 0,
            Top = 1,
            Text = text,
            Margin = Padding.Empty,
            FlatStyle = FlatStyle.System,
            TabStop = false
        };
    }

    private void ToolbarPanel_Resize(object? sender, EventArgs e)
    {
        LayoutToolbar();
    }

    private void LayoutToolbar()
    {
        var controlHeight = Math.Max(24, _toolbarPanel.ClientSize.Height - 2);
        var top = Math.Max(0, (_toolbarPanel.ClientSize.Height - controlHeight) / 2);

        _toggleTreeExpansionButton.SetBounds(0, top, 30, controlHeight);
        _toggleTreePanelButton.SetBounds(_toggleTreeExpansionButton.Right + 6, top, 30, controlHeight);
        _searchTextBox.SetBounds(_toggleTreePanelButton.Right + 8, top, Math.Max(120, _toolbarPanel.ClientSize.Width - (_toggleTreePanelButton.Right + 12)), controlHeight);
    }

    private TreeNode BuildTreeNode(ProjectPaletteTreeNode sourceNode)
    {
        var visibleCount = _model is null
            ? sourceNode.DescendantItems.Count
            : sourceNode.DescendantItems.Count(item => IsDisplayableRightPanelItem(item, _model.Context));

        var suffix = visibleCount > 0
            ? $" ({visibleCount})"
            : string.Empty;

        var treeNode = new TreeNode($"{sourceNode.DisplayName}{suffix}")
        {
            Tag = sourceNode
        };

        foreach (var child in sourceNode.Children)
        {
            treeNode.Nodes.Add(BuildTreeNode(child));
        }

        return treeNode;
    }

    private void ToggleTreeExpansionButton_Click(object? sender, EventArgs e)
    {
        if (_treeExpanded)
        {
            _treeExpanded = false;
            CollapseTreeToRootElementsOnly();
        }
        else
        {
            _treeExpanded = true;
            ExpandTree();
        }

        UpdateToolbarButtons();
    }

    private void ToggleTreePanelButton_Click(object? sender, EventArgs e)
    {
        if (_treePanelVisible)
        {
            if (!_splitContainer.Panel1Collapsed)
            {
                _savedSplitterDistance = _splitContainer.SplitterDistance;
            }

            _splitContainer.Panel1Collapsed = true;
            _treePanelVisible = false;
        }
        else
        {
            _splitContainer.Panel1Collapsed = false;
            TryRestoreSplitterDistance();
            _treePanelVisible = true;
        }

        UpdateToolbarButtons();
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressSearchTextChanged || _model is null)
        {
            return;
        }

        var nowActive = HasActiveSearch;
        if (nowActive && !_searchActive)
        {
            _treeExpandedBeforeSearch = _treeExpanded;
            _treeExpanded = true;
            _searchActive = true;
        }
        else if (!nowActive && _searchActive)
        {
            _searchActive = false;
            _treeExpanded = _treeExpandedBeforeSearch;
        }

        ApplyTreeFilter(preserveSelection: true);
        UpdateToolbarButtons();
    }

    private bool HasActiveSearch => !string.IsNullOrWhiteSpace(_searchTextBox.Text);

    private void ApplyTreeFilter(bool preserveSelection)
    {
        if (_model is null)
        {
            return;
        }

        var selectedNodeKey = preserveSelection
            ? _selectedNodeKey ?? (_classTree.SelectedNode?.Tag as ProjectPaletteTreeNode)?.NodeKey
            : null;

        var rootNodes = GetVisibleRootNodes().ToList();

        _classTree.BeginUpdate();
        _classTree.Nodes.Clear();
        foreach (var root in rootNodes)
        {
            _classTree.Nodes.Add(BuildTreeNode(root));
        }
        _classTree.EndUpdate();

        if (_treeExpanded)
        {
            ExpandTree();
        }
        else
        {
            CollapseTreeToRootElementsOnly();
        }

        if (_classTree.Nodes.Count == 0)
        {
            _selectedNodeKey = null;
            RefreshItems(null);
            _statusLabel.Text = HasActiveSearch
                ? "No classes match the search."
                : "No class selected.";
            return;
        }

        if (!string.IsNullOrWhiteSpace(selectedNodeKey) && TrySelectNodeByKey(selectedNodeKey))
        {
            return;
        }

        _classTree.SelectedNode = _classTree.Nodes[0];
    }

    private IEnumerable<ProjectPaletteTreeNode> GetVisibleRootNodes()
    {
        if (_model is null)
        {
            return Array.Empty<ProjectPaletteTreeNode>();
        }

        if (!HasActiveSearch)
        {
            return _model.RootNodes;
        }

        var search = _searchTextBox.Text;
        var filtered = new List<ProjectPaletteTreeNode>();
        foreach (var root in _model.RootNodes)
        {
            var clone = FilterNode(root, search);
            if (clone is not null)
            {
                filtered.Add(clone);
            }
        }

        return filtered;
    }

    private static ProjectPaletteTreeNode? FilterNode(ProjectPaletteTreeNode sourceNode, string? searchText)
    {
        var filteredItems = sourceNode.DescendantItems
            .Where(item => ItemMatchesSearch(item, searchText))
            .ToList();

        var filteredChildren = sourceNode.Children
            .Select(child => FilterNode(child, searchText))
            .Where(child => child is not null)
            .Cast<ProjectPaletteTreeNode>()
            .ToList();

        var nodeMatches = NodeMatchesSearch(sourceNode, searchText);
        if (!nodeMatches && filteredItems.Count == 0 && filteredChildren.Count == 0)
        {
            return null;
        }

        var clone = new ProjectPaletteTreeNode(sourceNode.NodeKey, sourceNode.Category, sourceNode.DisplayName, sourceNode.ClassName, sourceNode.IsSynthetic)
        {
            PaletteItem = sourceNode.PaletteItem is not null && filteredItems.Any(item => string.Equals(item.UniqueKey, sourceNode.PaletteItem.UniqueKey, StringComparison.OrdinalIgnoreCase))
                ? sourceNode.PaletteItem
                : null,
            DescendantItems = filteredItems
        };

        clone.Children.AddRange(filteredChildren);
        return clone;
    }

    private static bool ItemMatchesSearch(ProjectPaletteItem item, string? searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return true;
        }

        var haystack = SearchText.Normalize(string.Join(' ', new[]
        {
            item.DisplayName,
            item.ClassName,
            item.VisualName,
            item.SymbolName,
            item.LineStyleName,
            item.ParentClassName,
            string.Join(' ', item.StyleCandidates)
        }.Where(value => !string.IsNullOrWhiteSpace(value))));

        return SearchText.Matches(haystack, searchText);
    }

    private static bool NodeMatchesSearch(ProjectPaletteTreeNode node, string? searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return true;
        }

        var haystack = SearchText.Normalize(string.Join(' ', new[]
        {
            node.DisplayName,
            node.ClassName,
            node.PaletteItem?.DisplayName,
            node.PaletteItem?.ClassName,
            node.PaletteItem?.VisualName
        }.Where(value => !string.IsNullOrWhiteSpace(value))));

        return SearchText.Matches(haystack, searchText);
    }

    private bool TrySelectNodeByKey(string nodeKey)
    {
        var treeNode = FindNodeByKey(_classTree.Nodes, nodeKey);
        if (treeNode is null)
        {
            return false;
        }

        _classTree.SelectedNode = treeNode;
        treeNode.EnsureVisible();
        return true;
    }

    private static TreeNode? FindNodeByKey(TreeNodeCollection nodes, string nodeKey)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Tag is ProjectPaletteTreeNode paletteTreeNode
                && string.Equals(paletteTreeNode.NodeKey, nodeKey, StringComparison.OrdinalIgnoreCase))
            {
                return node;
            }

            var match = FindNodeByKey(node.Nodes, nodeKey);
            if (match is not null)
            {
                return match;
            }
        }

        return null;
    }

    private void ExpandTree()
    {
        _classTree.BeginUpdate();
        _classTree.ExpandAll();
        _classTree.EndUpdate();
        _classTree.SelectedNode?.EnsureVisible();
    }

    private void CollapseTreeToRootElementsOnly()
    {
        _classTree.BeginUpdate();
        foreach (TreeNode root in _classTree.Nodes)
        {
            CollapseNodeRecursive(root);
        }
        _classTree.EndUpdate();
        _classTree.SelectedNode?.EnsureVisible();
    }

    private static void CollapseNodeRecursive(TreeNode node)
    {
        foreach (TreeNode child in node.Nodes)
        {
            CollapseNodeRecursive(child);
        }

        node.Collapse();
    }

    private void UpdateToolbarButtons()
    {
        _toggleTreeExpansionButton.Text = _treeExpanded ? "-" : "+";
        _toolTip.SetToolTip(
            _toggleTreeExpansionButton,
            _treeExpanded
                ? "Collapse the class tree to the root elements only."
                : "Expand the complete class tree.");

        _toggleTreePanelButton.Text = _treePanelVisible ? "<" : ">";
        _toolTip.SetToolTip(
            _toggleTreePanelButton,
            _treePanelVisible
                ? "Hide the class tree panel."
                : "Show the class tree panel.");
    }

    private void TryRestoreSplitterDistance()
    {
        try
        {
            _splitContainer.SplitterDistance = Math.Max(140, _savedSplitterDistance);
        }
        catch
        {
        }
    }

    private void ClassTree_AfterSelect(object? sender, TreeViewEventArgs e)
    {
        if (e.Node?.Tag is ProjectPaletteTreeNode selectedNode)
        {
            _selectedNodeKey = selectedNode.NodeKey;
        }
        else
        {
            _selectedNodeKey = null;
        }

        RefreshItems(e.Node?.Tag as ProjectPaletteTreeNode);
    }

    private void RefreshItems(ProjectPaletteTreeNode? selectedNode)
    {
        _itemList.BeginUpdate();
        _itemList.Items.Clear();
        _imageList.Images.Clear();

        if (_model is null || selectedNode is null)
        {
            _statusLabel.Text = HasActiveSearch
                ? "No classes match the search."
                : "No class selected.";
            _itemList.EndUpdate();
            return;
        }

        var insertableItems = selectedNode.DescendantItems
            .Where(item => IsDisplayableRightPanelItem(item, _model.Context))
            .ToList();

        foreach (var item in insertableItems)
        {
            var imageKey = item.UniqueKey;
            if (!_imageList.Images.ContainsKey(imageKey))
            {
                try
                {
                    _imageList.Images.Add(imageKey, _previewService.GetPreview(item, _model.Context));
                }
                catch (System.Exception ex)
                {
                    _imageList.Images.Add(imageKey, CreateEmergencyPreviewPlaceholder());
                    System.Diagnostics.Debug.WriteLine($"Preview creation failed for {item.ClassName}: {ex.Message}");
                }
            }

            var listViewItem = new ListViewItem(item.DisplayName)
            {
                ImageKey = imageKey,
                Tag = item,
                ToolTipText = BuildTooltip(item, _model.Context)
            };

            _itemList.Items.Add(listViewItem);
        }

        _statusLabel.Text = $"{BuildNodePath(_classTree.SelectedNode)}: {insertableItems.Count} insertable class(es)";
        _itemList.EndUpdate();
    }

    private void ItemList_MouseUp(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left || _model is null)
        {
            return;
        }

        var hit = _itemList.HitTest(e.Location);
        if (hit.Item?.Tag is not ProjectPaletteItem paletteItem)
        {
            return;
        }

        if (IsHandleCreated)
        {
            BeginInvoke(new Action(() => _executionService.Execute(paletteItem, _model.Context)));
        }
        else
        {
            _executionService.Execute(paletteItem, _model.Context);
        }
    }

    private string BuildTooltip(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var styleResolution = _previewService.GetStyleResolution(item, context);

        var parts = new List<string>
        {
            item.DisplayName,
            $"Class: {item.ClassName}",
            $"Block: {styleResolution.DisplayBlockName ?? "<unresolved>"}"
        };

        if (!string.IsNullOrWhiteSpace(styleResolution.StyleInfo?.StyleName))
        {
            parts.Add($"Style: {styleResolution.StyleInfo.StyleName}");
        }

        if (item.TpIncluded.HasValue)
        {
            parts.Add($"TPIncluded: {item.TpIncluded.Value}");
        }

        if (item.SupportedStandardsMask is int standards)
        {
            parts.Add($"SupportedStandards: {standards}");
        }

        return string.Join(Environment.NewLine, parts);
    }

    private bool IsDisplayableRightPanelItem(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        return item.IsLeafClass
            && item.StyleCandidates.Count > 0;
    }

    private static Bitmap CreateEmergencyPreviewPlaceholder()
    {
        var bitmap = new Bitmap(SymbolPreviewService.PreviewPixelSize, SymbolPreviewService.PreviewPixelSize);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(SystemColors.Window);
        using var borderPen = new Pen(SystemColors.ControlDark);
        graphics.DrawRectangle(borderPen, 0, 0, bitmap.Width - 1, bitmap.Height - 1);
        using var pen = new Pen(SystemColors.ControlText, 2f);
        graphics.DrawRectangle(pen, 18, 18, bitmap.Width - 36, bitmap.Height - 36);
        graphics.DrawLine(pen, 18, bitmap.Height / 2f, bitmap.Width - 18, bitmap.Height / 2f);
        return bitmap;
    }

    private static string BuildNodePath(TreeNode? treeNode)
    {
        if (treeNode is null)
        {
            return string.Empty;
        }

        var parts = new Stack<string>();
        var current = treeNode;
        while (current is not null)
        {
            parts.Push(current.Text.Split('(')[0].Trim());
            current = current.Parent;
        }

        return string.Join(" > ", parts);
    }

    private sealed class BufferedTreeView : TreeView
    {
        public BufferedTreeView()
        {
            DrawMode = TreeViewDrawMode.Normal;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            UpdateStyles();
        }
    }

    private sealed class BufferedListView : ListView
    {
        public BufferedListView()
        {
            DoubleBuffered = true;
        }
    }
}
