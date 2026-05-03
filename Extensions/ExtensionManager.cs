using static ExcelConsole.Extensions.ExtensionProtocol;

namespace ExcelConsole.Extensions;

/// <summary>
/// Orchestrates extension lifecycle: scans grid for ext:/prefix: cells,
/// installs from GitHub, launches processes, routes messages, writes cell output.
/// </summary>
public class ExtensionManager : IDisposable
{
    private readonly IExtensionEnvironment _env;
    private GridManager _grid;

    // Extension source (from ext: cell) -> installation state
    private readonly Dictionary<string, LoadedExtension> _extensions = new();
    // Registered prefix -> extension name
    private readonly Dictionary<string, string> _prefixMap = new();
    // Tracks which ext: cells we've already processed
    private readonly HashSet<(int row, int col)> _processedExtCells = new();
    // Tracks active prefix calls: anchor cell -> activation state
    private readonly Dictionary<(int row, int col), ActiveCall> _activeCalls = new();

    private bool _disposed;

    /// <summary>
    /// True if any extension wrote cells since last check. Reset by calling ConsumeHasChanges().
    /// </summary>
    public bool HasChanges { get; private set; }

    public ExtensionManager(IExtensionEnvironment env, GridManager grid)
    {
        _env = env;
        _grid = grid;
    }

    /// <summary>
    /// Returns and resets the HasChanges flag. Used by the UI to know when to repaint.
    /// </summary>
    public bool ConsumeHasChanges()
    {
        bool val = HasChanges;
        HasChanges = false;
        return val;
    }

    /// <summary>
    /// The set of currently registered extension prefixes.
    /// </summary>
    public HashSet<string> KnownPrefixes => new(_prefixMap.Keys);

    /// <summary>
    /// Returns the status of a cell for color coding purposes.
    /// </summary>
    public ExtensionCellStatus GetCellStatus(int row, int col)
    {
        string val = _grid.GetCellValue(row, col);

        // Check if it's an ext: cell
        if (CellPrefix.IsExtension(val))
        {
            string? source = CellPrefix.ParseExtensionSource(val);
            if (source == null) return ExtensionCellStatus.None;

            if (val.Contains("[install failed]") || val.Contains("[bad manifest]") || val.Contains("[start failed]"))
                return ExtensionCellStatus.Error;

            if (_extensions.TryGetValue(source, out var ext) && ext.Process.IsRunning)
                return ExtensionCellStatus.Running;

            return ExtensionCellStatus.Loading;
        }

        // Check if it's a prefix call cell
        if (_activeCalls.ContainsKey((row, col)))
        {
            var call = _activeCalls[(row, col)];
            if (call.HasError)
                return ExtensionCellStatus.Error;
            return ExtensionCellStatus.Running;
        }

        return ExtensionCellStatus.None;
    }

    /// <summary>
    /// Forces re-activation of an extension call at the given cell.
    /// Removes any existing activation so the next ScanGrid() will re-send activate.
    /// </summary>
    public void ReactivateCell(int row, int col)
    {
        _activeCalls.Remove((row, col));
    }

    /// <summary>
    /// Updates the grid reference after a rebuild. Clears activation tracking
    /// so extensions are re-scanned against the new grid.
    /// </summary>
    public void UpdateGrid(GridManager newGrid)
    {
        _grid = newGrid;
        _activeCalls.Clear();
        _processedExtCells.Clear();
    }

    public void ScanGrid()
    {
        if (_disposed) return;

        // Phase 1: Find and process ext: cells
        for (int r = 0; r < _grid.RowCount; r++)
        {
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                string val = _grid.GetCellValue(r, c);
                if (CellPrefix.IsExtension(val))
                {
                    if (!_processedExtCells.Contains((r, c)))
                    {
                        ProcessExtCell(r, c, val);
                        _processedExtCells.Add((r, c));
                    }
                }
            }
        }

        // Phase 2: Process messages from running extensions
        ProcessIncomingMessages();

        // Phase 3: Find and route prefix: cells
        var knownPrefixes = KnownPrefixes;
        if (knownPrefixes.Count == 0) return;

        for (int r = 0; r < _grid.RowCount; r++)
        {
            for (int c = 0; c < _grid.ColumnCount; c++)
            {
                string val = _grid.GetCellValue(r, c);
                if (string.IsNullOrEmpty(val)) continue;

                var parsed = CellPrefix.ParseExtensionCall(val, knownPrefixes);
                if (parsed == null) continue;

                var anchor = (r, c);
                if (_activeCalls.ContainsKey(anchor)) continue; // already activated

                ActivateExtensionCall(r, c, parsed.Value);
            }
        }
    }

    /// <summary>
    /// Processes an ext: cell. Installs the extension from GitHub if needed, then launches it.
    /// </summary>
    private void ProcessExtCell(int row, int col, string cellValue)
    {
        string? source = CellPrefix.ParseExtensionSource(cellValue);
        if (source == null) return;

        if (_extensions.ContainsKey(source)) return; // already loaded

        // Install from GitHub
        string? extDir = ExtensionInstaller.Install(source, _env);
        if (extDir == null)
        {
            _grid.SetCellValue(row, col, $"ext: {source} [install failed]");
            return;
        }

        var manifest = ExtensionInstaller.ReadManifest(extDir);
        if (manifest == null)
        {
            _grid.SetCellValue(row, col, $"ext: {source} [bad manifest]");
            return;
        }

        // Launch process
        var process = new ExtensionProcess(manifest.Name);
        bool started = process.Start(_env, manifest.Entry, extDir);
        if (!started)
        {
            _grid.SetCellValue(row, col, $"ext: {source} [start failed]");
            process.Dispose();
            return;
        }

        _extensions[source] = new LoadedExtension
        {
            Source = source,
            Manifest = manifest,
            Process = process,
            Directory = extDir
        };
    }

    /// <summary>
    /// Drains incoming messages from all running extensions and handles them.
    /// </summary>
    private void ProcessIncomingMessages()
    {
        foreach (var ext in _extensions.Values)
        {
            if (!ext.Process.IsRunning) continue;

            var messages = ext.Process.DrainMessages();
            foreach (var json in messages)
            {
                string? type = GetMessageType(json);
                switch (type)
                {
                    case "register":
                        var reg = Deserialize<RegisterMessage>(json);
                        if (reg != null)
                        {
                            ext.Process.HandleRegister(reg);
                            string prefix = reg.Prefix.ToLowerInvariant();
                            _prefixMap.TryAdd(prefix, ext.Source);
                        }
                        break;

                    case "write":
                        var write = Deserialize<WriteCellsMessage>(json);
                        if (write != null)
                            HandleWriteCells(write);
                        break;

                    case "status":
                        var status = Deserialize<StatusMessage>(json);
                        if (status != null)
                            HandleStatus(status);
                        break;

                    case "error":
                        var error = Deserialize<ErrorMessage>(json);
                        if (error != null)
                            HandleError(error);
                        break;

                    case "log":
                        // Could log to debug output in the future
                        break;
                }
            }
        }
    }

    /// <summary>
    /// Sends an activate message to the extension that owns the given prefix.
    /// </summary>
    private void ActivateExtensionCall(int row, int col, (string prefix, string[] extParams, int gridCols, int gridRows) call)
    {
        if (!_prefixMap.TryGetValue(call.prefix, out string? source)) return;
        if (!_extensions.TryGetValue(source, out var ext)) return;
        if (!ext.Process.IsRunning) return;

        string activationId = $"activate-{row}-{col}";

        // Output starts one row below the command cell
        var anchor = new CellPosition { Row = row + 1, Col = col };

        // Expand cell references (e.g., {A1::C10}) in params before sending
        var expandedParams = new string[call.extParams.Length];
        for (int i = 0; i < call.extParams.Length; i++)
            expandedParams[i] = CellPrefix.ExpandCellReferences(call.extParams[i], _grid);

        var msg = new ActivateMessage
        {
            Id = activationId,
            Anchor = anchor,
            Params = expandedParams,
            GridCols = call.gridCols,
            GridRows = call.gridRows
        };

        ext.Process.SendMessage(msg);

        _activeCalls[(row, col)] = new ActiveCall
        {
            ActivationId = activationId,
            ExtensionSource = source,
            AnchorRow = row + 1,
            AnchorCol = col,
            GridCols = call.gridCols,
            GridRows = call.gridRows
        };
    }

    /// <summary>
    /// Handles a write-cells message from an extension, writing values into the grid.
    /// </summary>
    private void HandleWriteCells(WriteCellsMessage msg)
    {
        // Find the active call for this message
        ActiveCall? call = null;
        foreach (var ac in _activeCalls.Values)
        {
            if (ac.ActivationId == msg.Id)
            {
                call = ac;
                break;
            }
        }

        if (call == null) return;

        foreach (var cell in msg.Cells)
        {
            int targetRow = call.AnchorRow + cell.Row;
            int targetCol = call.AnchorCol + cell.Col;

            if (targetRow >= 0 && targetRow < _grid.RowCount &&
                targetCol >= 0 && targetCol < _grid.ColumnCount)
            {
                _grid.SetCellValue(targetRow, targetCol, cell.Value);
            }
        }

        HasChanges = true;
    }

    /// <summary>
    /// Handles a status message from an extension, writing temporary status text to the anchor cell.
    /// </summary>
    private void HandleStatus(StatusMessage msg)
    {
        foreach (var ac in _activeCalls.Values)
        {
            if (ac.ActivationId == msg.Id)
            {
                if (ac.AnchorRow >= 0 && ac.AnchorRow < _grid.RowCount &&
                    ac.AnchorCol >= 0 && ac.AnchorCol < _grid.ColumnCount)
                {
                    _grid.SetCellValue(ac.AnchorRow, ac.AnchorCol, msg.Message);
                    HasChanges = true;
                }
                break;
            }
        }
    }

    /// <summary>
    /// Handles an error message from an extension.
    /// </summary>
    private void HandleError(ErrorMessage msg)
    {
        // Find the active call and write error to the anchor cell
        foreach (var ac in _activeCalls.Values)
        {
            if (ac.ActivationId == msg.Id)
            {
                ac.HasError = true;
                if (ac.AnchorRow >= 0 && ac.AnchorRow < _grid.RowCount &&
                    ac.AnchorCol >= 0 && ac.AnchorCol < _grid.ColumnCount)
                {
                    _grid.SetCellValue(ac.AnchorRow, ac.AnchorCol, $"[err: {msg.Message}]");
                    HasChanges = true;
                }
                break;
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        foreach (var ext in _extensions.Values)
            ext.Process.Dispose();
        _extensions.Clear();
        _prefixMap.Clear();
        _activeCalls.Clear();
    }

    // ── Internal types ───────────────────────────────────────────────

    private class LoadedExtension
    {
        public string Source { get; set; } = "";
        public ExtensionManifest Manifest { get; set; } = new();
        public ExtensionProcess Process { get; set; } = null!;
        public string Directory { get; set; } = "";
    }

    private class ActiveCall
    {
        public string ActivationId { get; set; } = "";
        public string ExtensionSource { get; set; } = "";
        public int AnchorRow { get; set; }
        public int AnchorCol { get; set; }
        public int GridCols { get; set; }
        public int GridRows { get; set; }
        public bool HasError { get; set; }
    }
}

/// <summary>
/// Status of a cell for extension color coding.
/// </summary>
public enum ExtensionCellStatus
{
    None,
    Loading,
    Running,
    Error
}
