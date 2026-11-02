import { type ChangeEvent, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  ArrowDownUp,
  Check,
  CheckCircle2,
  Code2,
  FileSpreadsheet,
  FileUp,
  HardDrive,
  LockKeyhole,
  RotateCcw,
  Search,
  ShieldCheck,
  Sparkles,
  X,
} from "lucide-react";
import { Link } from "wouter";

type WorkbookData = { file: File; workbook: XLSX.WorkBook };
type Side = "before" | "after";
type DiffKind = "changed" | "added" | "removed";
type Difference = {
  address: string;
  kind: DiffKind;
  before: string;
  after: string;
  beforeRaw: unknown;
  afterRaw: unknown;
};

const EMPTY_SHEET: XLSX.WorkSheet = {};
const cellValue = (cell: XLSX.CellObject | undefined) => {
  if (!cell || (cell.v === undefined && !cell.f)) return "";
  if (cell.w !== undefined) return String(cell.w);
  if (cell.f) return `=${cell.f}`;
  return String(cell.v);
};
const rawCellValue = (cell: XLSX.CellObject | undefined) => {
  if (!cell || (cell.v === undefined && !cell.f)) return undefined;
  const value = cell.v instanceof Date ? { date: cell.v.getTime() } : cell.v;
  return cell.f
    ? JSON.stringify({ formula: cell.f, value })
    : JSON.stringify(value);
};
const isPopulated = (cell: XLSX.CellObject | undefined) =>
  !!cell && (cell.v !== undefined || !!cell.f);
const fileSize = (size: number) =>
  size < 1024 * 1024
    ? `${Math.max(1, Math.round(size / 1024))} KB`
    : `${(size / 1024 / 1024).toFixed(1)} MB`;

function App() {
  const [before, setBefore] = useState<WorkbookData | null>(null);
  const [after, setAfter] = useState<WorkbookData | null>(null);
  const [beforeSheet, setBeforeSheet] = useState("");
  const [afterSheet, setAfterSheet] = useState("");
  const [errors, setErrors] = useState<Record<Side, string>>({
    before: "",
    after: "",
  });
  const [dragging, setDragging] = useState<Side | null>(null);
  const [parsing, setParsing] = useState<Side | null>(null);
  const [result, setResult] = useState<{
    differences: Difference[];
    beforeSheet: string;
    afterSheet: string;
  } | null>(null);
  const [notice, setNotice] = useState("");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const [selectedAddress, setSelectedAddress] = useState("");
  const beforeInput = useRef<HTMLInputElement>(null);
  const afterInput = useRef<HTMLInputElement>(null);

  const beforeNames = before?.workbook.SheetNames ?? [];
  const afterNames = after?.workbook.SheetNames ?? [];
  const canCompare =
    !!before && !!after && !!beforeSheet && !!afterSheet && !parsing;
  const beforeWorksheet =
    result && before
      ? before.workbook.Sheets[result.beforeSheet] ?? EMPTY_SHEET
      : EMPTY_SHEET;
  const afterWorksheet =
    result && after
      ? after.workbook.Sheets[result.afterSheet] ?? EMPTY_SHEET
      : EMPTY_SHEET;

  const compareSheets = (left: XLSX.WorkSheet, right: XLSX.WorkSheet) => {
    const addresses = new Set<string>();
    for (const key of [...Object.keys(left), ...Object.keys(right)])
      if (/^[A-Z]{1,3}\d+$/.test(key)) addresses.add(key);
    const differences: Difference[] = [];
    for (const address of addresses) {
      const l = left[address] as XLSX.CellObject | undefined;
      const r = right[address] as XLSX.CellObject | undefined;
      const leftHas = isPopulated(l),
        rightHas = isPopulated(r);
      if (!leftHas && !rightHas) continue;
      const leftValue = rawCellValue(l),
        rightValue = rawCellValue(r);
      if (leftHas && rightHas && Object.is(leftValue, rightValue)) continue;
      differences.push({
        address,
        kind: !leftHas ? "added" : !rightHas ? "removed" : "changed",
        before: cellValue(l),
        after: cellValue(r),
        beforeRaw: leftValue,
        afterRaw: rightValue,
      });
    }
    return differences.sort((a, b) => {
      const aa = XLSX.utils.decode_cell(a.address),
        bb = XLSX.utils.decode_cell(b.address);
      return aa.r - bb.r || aa.c - bb.c;
    });
  };

  const visibleDifferences = useMemo(() => {
    if (!result) return [];
    const query = search.trim().toLowerCase();
    return result.differences.filter(
      (item) =>
        (filter === "all" || item.kind === filter) &&
        (!query ||
          `${item.address} ${item.before} ${item.after} ${item.kind}`
            .toLowerCase()
            .includes(query))
    );
  }, [result, search, filter]);

  const differencesByAddress = useMemo(
    () =>
      new Map(
        (result?.differences ?? []).map((difference) => [
          difference.address,
          difference,
        ])
      ),
    [result]
  );

  const counts = useMemo(
    () => ({
      changed:
        result?.differences.filter((item) => item.kind === "changed").length ??
        0,
      added:
        result?.differences.filter((item) => item.kind === "added").length ?? 0,
      removed:
        result?.differences.filter((item) => item.kind === "removed").length ??
        0,
    }),
    [result]
  );

  const receiveFile = async (file: File, side: Side) => {
    setErrors((current) => ({ ...current, [side]: "" }));
    setNotice("");
    setResult(null);
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      setErrors((current) => ({
        ...current,
        [side]:
          "Vui lòng chọn file Excel đúng định dạng (.xlsx, .xlsm hoặc .xls).",
      }));
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      setErrors((current) => ({
        ...current,
        [side]:
          "Dung lượng file vượt quá giới hạn 50 MB. Vui lòng chọn file nhỏ hơn.",
      }));
      return;
    }
    setParsing(side);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
        cellNF: true,
      });
      if (!workbook.SheetNames?.length)
        throw new Error("File Excel này không chứa bất kỳ worksheet nào.");
      const data = { file, workbook };
      if (side === "before") {
        setBefore(data);
        setBeforeSheet(workbook.SheetNames[0]);
      } else {
        setAfter(data);
        setAfterSheet(workbook.SheetNames[0]);
      }
    } catch (error) {
      const message =
        error instanceof Error && error.message.includes("không chứa")
          ? error.message
          : "Không thể đọc file này. Vui lòng thử lưu lại file dưới dạng .xlsx.";
      setErrors((current) => ({ ...current, [side]: message }));
      if (side === "before") {
        setBefore(null);
        setBeforeSheet("");
      } else {
        setAfter(null);
        setAfterSheet("");
      }
    } finally {
      setParsing(null);
    }
  };

  const handleInput = (event: ChangeEvent<HTMLInputElement>, side: Side) => {
    const file = event.target.files?.[0];
    if (file) void receiveFile(file, side);
    event.target.value = "";
  };

  const removeFile = (side: Side) => {
    setErrors((current) => ({ ...current, [side]: "" }));
    setNotice("");
    setResult(null);
    if (side === "before") {
      setBefore(null);
      setBeforeSheet("");
    } else {
      setAfter(null);
      setAfterSheet("");
    }
  };

  const reset = () => {
    setBefore(null);
    setAfter(null);
    setBeforeSheet("");
    setAfterSheet("");
    setErrors({ before: "", after: "" });
    setResult(null);
    setNotice("");
    setSearch("");
    setFilter("all");
    setSelectedAddress("");
    if (beforeInput.current) beforeInput.current.value = "";
    if (afterInput.current) afterInput.current.value = "";
  };

  const runCompare = () => {
    if (!before || !after || !beforeSheet || !afterSheet) return;
    const differences = compareSheets(
      before.workbook.Sheets[beforeSheet] ?? EMPTY_SHEET,
      after.workbook.Sheets[afterSheet] ?? EMPTY_SHEET
    );
    setResult({ differences, beforeSheet, afterSheet });
    setSearch("");
    setFilter("all");
    setSelectedAddress("");
    setNotice(
      differences.length
        ? ""
        : "Hai trang tính này hoàn toàn giống nhau. Không tìm thấy sự khác biệt nào."
    );
  };

  const workbookOnlyNames = (side: Side) => {
    if (!before || !after) return [];
    const source = side === "before" ? beforeNames : afterNames;
    const other = side === "before" ? afterNames : beforeNames;
    return source.filter((name) => !other.includes(name));
  };
  const otherSideSheetMissing =
    result &&
    (!before?.workbook.Sheets[result.beforeSheet] ||
      !after?.workbook.Sheets[result.afterSheet]);

  const renderUploadCard = (side: Side) => {
    const isBefore = side === "before";
    const data = isBefore ? before : after;
    const title = isBefore ? "Bản gốc (File cũ)" : "Bản sửa đổi (File mới)";
    const version = isBefore ? "01 / FILE GỐC" : "02 / FILE MỚI";
    const inputRef = isBefore ? beforeInput : afterInput;
    return (
      <section
        className={`upload-card ${data ? "has-file" : ""} ${
          dragging === side ? "is-dragging" : ""
        }`}
        data-testid={`upload-card-${side}`}
        onDragOver={(event) => {
          event.preventDefault();
          setDragging(side);
        }}
        onDragLeave={(event) => {
          if (!event.currentTarget.contains(event.relatedTarget as Node))
            setDragging(null);
        }}
        onDrop={(event) => {
          event.preventDefault();
          setDragging(null);
          const file = event.dataTransfer.files[0];
          if (file) void receiveFile(file, side);
        }}
      >
        <div className="upload-heading">
          <div className="upload-label">
            <span className="file-icon">
              <FileSpreadsheet size={17} />
            </span>
            <strong>{title}</strong>
          </div>
          <span className="version-tag">{version}</span>
        </div>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xlsm,.xls"
          className="sr-only"
          onChange={(event) => handleInput(event, side)}
          data-testid={`input-workbook-${side}`}
          aria-label={`Chọn file ${title.toLowerCase()}`}
        />
        {data ? (
          <div className="drop-zone" data-testid={`file-loaded-${side}`}>
            <div className="file-detail">
              <span className="drop-symbol">
                <Check size={18} />
              </span>
              <div className="file-detail-copy">
                <strong title={data.file.name}>{data.file.name}</strong>
                <span>
                  {fileSize(data.file.size)} · {data.workbook.SheetNames.length}{" "}
                  {data.workbook.SheetNames.length === 1
                    ? "worksheet"
                    : "worksheets"}
                </span>
              </div>
              <button
                className="remove-file"
                onClick={() => removeFile(side)}
                aria-label={`Xóa file ${title.toLowerCase()}`}
                data-testid={`button-remove-${side}`}
              >
                <X size={16} />
              </button>
            </div>
          </div>
        ) : (
          <button
            type="button"
            className="drop-zone"
            onClick={() => inputRef.current?.click()}
            data-testid={`button-upload-${side}`}
          >
            <span className="drop-symbol">
              <FileUp size={18} />
            </span>
            <span className="drop-copy">
              <strong>
                {parsing === side
                  ? "Đang đọc dữ liệu file…"
                  : "Kéo thả file Excel vào đây, hoặc bấm để duyệt file"}
              </strong>
              <span>Hỗ trợ file .xlsx, .xlsm, .xls · Tối đa 50 MB</span>
            </span>
          </button>
        )}
        {errors[side] && (
          <p
            className="upload-error"
            role="alert"
            data-testid={`error-upload-${side}`}
          >
            {errors[side]}
          </p>
        )}
      </section>
    );
  };

  const workbookOnly = [
    ...new Set([...workbookOnlyNames("before"), ...workbookOnlyNames("after")]),
  ];
  const gridCellClass = (address: string) => {
    const diff = differencesByAddress.get(address);
    return `${diff?.kind ?? ""} ${
      selectedAddress === address ? "focus-cell" : ""
    }`.trim();
  };
  const rangeRef = (sheet: XLSX.WorkSheet) => {
    try {
      return sheet["!ref"]
        ? XLSX.utils.decode_range(sheet["!ref"])
        : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    } catch {
      return { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    }
  };
  const fullGridRows = result
    ? Math.max(rangeRef(beforeWorksheet).e.r, rangeRef(afterWorksheet).e.r, 0) +
      1
    : 0;
  const fullGridColumns = result
    ? Math.max(rangeRef(beforeWorksheet).e.c, rangeRef(afterWorksheet).e.c, 0) +
      1
    : 0;
  const gridRows = result ? Math.min(35, fullGridRows) : 0;
  const gridColumns = result ? Math.min(12, fullGridColumns) : 0;
  const selectedCoordinates = selectedAddress
    ? XLSX.utils.decode_cell(selectedAddress)
    : null;
  const gridStartRow =
    result && selectedCoordinates && fullGridRows > gridRows
      ? Math.min(
          Math.max(selectedCoordinates.r - Math.floor(gridRows / 2), 0),
          fullGridRows - gridRows
        )
      : 0;
  const gridStartColumn =
    result && selectedCoordinates && fullGridColumns > gridColumns
      ? Math.min(
          Math.max(selectedCoordinates.c - Math.floor(gridColumns / 2), 0),
          fullGridColumns - gridColumns
        )
      : 0;
  const gridCols = Array.from({ length: gridColumns }, (_, index) =>
    XLSX.utils.encode_col(gridStartColumn + index)
  );
  const grid = (side: Side, sheet: XLSX.WorkSheet) => (
    <div className="grid-wrap" data-testid={`grid-${side}`}>
      <table
        className="comparison-grid"
        aria-label={`Bảng dữ liệu ${
          side === "before" ? "File cũ" : "File mới"
        }`}
      >
        <thead>
          <tr>
            <th className="row-header"></th>
            {gridCols.map((column) => (
              <th key={column}>{column}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: gridRows }, (_, row) => {
            const rowIndex = gridStartRow + row;
            return (
              <tr key={rowIndex}>
                <th className="row-header">{rowIndex + 1}</th>
                {gridCols.map((column) => {
                  const address = `${column}${rowIndex + 1}`;
                  const cell = sheet[address] as XLSX.CellObject | undefined;
                  return (
                    <td
                      className={gridCellClass(address)}
                      key={address}
                      title={`${address}${
                        cellValue(cell) ? ` · ${cellValue(cell)}` : ""
                      }`}
                      data-testid={`cell-${side}-${address}`}
                    >
                      {cellValue(cell)}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );

  const translateKind = (kind: DiffKind) => {
    switch (kind) {
      case "changed":
        return "Thay đổi";
      case "added":
        return "Thêm mới";
      case "removed":
        return "Bị xóa";
    }
  };

  return (
    <div className="app-shell">
      <header className="topbar">
        <a
          href="/"
          className="brand"
          data-testid="link-home"
          aria-label="Trang chủ Sheetwise"
        >
          <span className="brand-mark">
            <ArrowDownUp size={18} strokeWidth={2.2} />
          </span>
          <span className="brand-name">
            sheetwise<span> / so sánh excel</span>
          </span>
        </a>
        <div className="topbar-tools">
          <Link
            href="/python"
            className="python-nav-link"
            data-testid="link-python-runner"
          >
            <Code2 size={15} /> Trình chạy Python
          </Link>
          <div className="topbar-note">
            <ShieldCheck size={15} /> Dữ liệu được bảo mật trên máy tính của bạn
          </div>
        </div>
      </header>

      <main className="workspace">
        <section className="intro">
          <div>
            <div className="eyebrow">
              <span className="eyebrow-line" /> ĐỐI CHIẾU DỮ LIỆU BẢNG TÍNH
            </div>
            <h1>Xem những gì đã thay đổi.</h1>
            <p className="intro-copy">
              So sánh hai file Excel song song. Phát hiện chính xác từng giá trị
              bị sửa đổi, ô mới thêm hoặc ô bị xóa nhanh chóng.
            </p>
          </div>
          <div className="privacy-note">
            <LockKeyhole size={15} /> Bảo mật tuyệt đối. Dữ liệu xử lý trực tiếp
            trên trình duyệt.
          </div>
        </section>

        <section className="upload-panel" aria-label="Chọn file và sheet Excel">
          <div className="upload-grid">
            {renderUploadCard("before")}
            <div className="versus" aria-hidden="true">
              VS
            </div>
            {renderUploadCard("after")}
          </div>

          <div className="sheet-row">
            <div className="sheet-field">
              <label htmlFor="sheet-before">SHEET FILE CŨ (BẢN GỐC)</label>
              <select
                id="sheet-before"
                value={beforeSheet}
                disabled={!before || parsing !== null}
                onChange={(event) => {
                  setBeforeSheet(event.target.value);
                  setResult(null);
                  setNotice("");
                }}
                data-testid="select-sheet-before"
              >
                {!beforeNames.length && (
                  <option value="">Chọn file Excel trước</option>
                )}
                {beforeNames.map((name) => (
                  <option value={name} key={name}>
                    {name}
                  </option>
                ))}
              </select>
            </div>
            <div className="sheet-field">
              <label htmlFor="sheet-after">SHEET FILE MỚI (BẢN SỬA ĐỔI)</label>
              <select
                id="sheet-after"
                value={afterSheet}
                disabled={!after || parsing !== null}
                onChange={(event) => {
                  setAfterSheet(event.target.value);
                  setResult(null);
                  setNotice("");
                }}
                data-testid="select-sheet-after"
              >
                {!afterNames.length && (
                  <option value="">Chọn file Excel trước</option>
                )}
                {afterNames.map((name) => (
                  <option value={name} key={name}>
                    {name}
                  </option>
                ))}
              </select>
            </div>
          </div>
          {workbookOnly.length > 0 && (
            <div
              className="sheet-missing"
              role="status"
              data-testid="status-worksheet-mismatch"
            >
              Worksheet{workbookOnly.length > 1 ? "s" : ""} chỉ có ở 1 file:{" "}
              {workbookOnly.map((name) => `“${name}”`).join(", ")}. Hãy chọn các
              cặp sheet tương ứng ở trên để so sánh.
            </div>
          )}
          <div className="action-row">
            <div className="local-hint">
              <HardDrive size={14} /> File Excel được đọc trực tiếp tại máy và
              không bị tải lên server.
            </div>
            <div className="action-buttons">
              <button
                className="btn btn-quiet"
                type="button"
                onClick={reset}
                data-testid="button-reset"
              >
                <RotateCcw size={14} /> Làm mới
              </button>
              <button
                className="btn btn-primary"
                type="button"
                onClick={runCompare}
                disabled={!canCompare}
                data-testid="button-compare"
              >
                <ArrowDownUp size={15} /> Bắt đầu so sánh
              </button>
            </div>
          </div>
        </section>

        {notice && (
          <div
            className={`status-banner ${
              notice.includes("giống nhau") ? "" : "warning"
            }`}
            role="status"
            data-testid="status-compare"
          >
            {notice.includes("giống nhau") ? (
              <CheckCircle2 size={18} />
            ) : (
              <Sparkles size={18} />
            )}
            <span>{notice}</span>
          </div>
        )}

        {result && (
          <section
            className="results-section"
            aria-label="Kết quả so sánh Excel"
            data-testid="section-results"
          >
            <div className="results-heading">
              <div>
                <div className="eyebrow">
                  <span className="eyebrow-line" /> HOÀN THÀNH SO SÁNH
                </div>
                <h2>
                  {result.differences.length
                    ? "Chi tiết các điểm khác biệt giữa 2 file."
                    : "Hai file hoàn toàn giống nhau."}
                </h2>
                <div className="results-caption">
                  Đang đối chiếu sheet <strong>{result.beforeSheet}</strong> với
                  sheet <strong>{result.afterSheet}</strong> theo tọa độ từng ô
                </div>
              </div>
              <div className="result-count">
                <strong>{result.differences.length}</strong> điểm khác biệt
              </div>
            </div>

            {result.differences.length > 0 ? (
              <>
                <div
                  className="summary-strip"
                  data-testid="summary-difference-counts"
                >
                  {(["changed", "added", "removed"] as DiffKind[]).map(
                    (kind) => (
                      <div
                        className="summary-stat"
                        key={kind}
                        data-testid={`summary-${kind}`}
                      >
                        <span className={`stat-dot ${kind}`} />
                        <div className="stat-copy">
                          <span>{translateKind(kind)}</span>
                          <strong>{counts[kind]}</strong>
                        </div>
                      </div>
                    )
                  )}
                </div>
                <div className="results-layout">
                  <section
                    className="diff-panel"
                    aria-label="Danh sách điểm khác biệt"
                  >
                    <div className="panel-head">
                      <div className="panel-headline">
                        <h3>Danh sách khác biệt</h3>
                        <span className="panel-subtitle">
                          {visibleDifferences.length} /{" "}
                          {result.differences.length} ô
                        </span>
                      </div>
                      <div className="diff-tools">
                        <div className="search-box">
                          <Search size={14} />
                          <input
                            type="search"
                            placeholder="Tìm ô hoặc giá trị..."
                            value={search}
                            onChange={(event) => setSearch(event.target.value)}
                            aria-label="Tìm kiếm điểm khác biệt"
                            data-testid="input-search-differences"
                          />
                        </div>
                        <select
                          className="filter-select"
                          value={filter}
                          onChange={(event) => setFilter(event.target.value)}
                          aria-label="Lọc loại khác biệt"
                          data-testid="select-filter-differences"
                        >
                          <option value="all">Tất cả loại</option>
                          <option value="changed">Thay đổi</option>
                          <option value="added">Thêm mới</option>
                          <option value="removed">Bị xóa</option>
                        </select>
                      </div>
                    </div>
                    <div className="diff-list" data-testid="list-differences">
                      {visibleDifferences.length ? (
                        visibleDifferences.map((difference) => (
                          <button
                            type="button"
                            className={`diff-item ${
                              selectedAddress === difference.address
                                ? "selected"
                                : ""
                            }`}
                            key={difference.address}
                            onClick={() =>
                              setSelectedAddress(difference.address)
                            }
                            data-testid={`difference-${difference.address}`}
                          >
                            <div className="diff-item-top">
                              <span className="diff-address">
                                Ô {difference.address}
                              </span>
                              <span className={`type-pill ${difference.kind}`}>
                                {translateKind(difference.kind)}
                              </span>
                            </div>
                            <div className="diff-values">
                              <div className="value-block before">
                                <small>Giá trị cũ</small>
                                <span>{difference.before || "—"}</span>
                              </div>
                              <div className="value-block after">
                                <small>Giá trị mới</small>
                                <span>{difference.after || "—"}</span>
                              </div>
                            </div>
                          </button>
                        ))
                      ) : (
                        <div
                          className="list-empty"
                          data-testid="empty-filtered-differences"
                        >
                          <Search size={18} />
                          <br />
                          Không có ô nào khớp với từ khóa tìm kiếm.
                          <br />
                          Hãy thử đổi từ khóa hoặc bộ lọc.
                        </div>
                      )}
                    </div>
                  </section>

                  <section
                    className="sheet-panel"
                    aria-label="Giao diện xem bảng tính"
                  >
                    <div className="sheet-head">
                      <div className="sheet-identity">
                        <span className="file-icon">
                          <FileSpreadsheet size={16} />
                        </span>
                        <div>
                          <strong>
                            {result.beforeSheet}{" "}
                            <span aria-hidden="true">↔</span>{" "}
                            {result.afterSheet}
                          </strong>
                          <span>Xem song song từng ô dữ liệu</span>
                        </div>
                      </div>
                      <div
                        className="grid-legend"
                        aria-label="Chú thích màu sắc ô"
                      >
                        <span className="legend-item">
                          <i className="legend-swatch changed" /> Thay đổi
                        </span>
                        <span className="legend-item">
                          <i className="legend-swatch added" /> Thêm mới
                        </span>
                        <span className="legend-item">
                          <i className="legend-swatch removed" /> Bị xóa
                        </span>
                      </div>
                    </div>
                    {workbookOnly.length > 0 && (
                      <div
                        className="only-sheet-callout"
                        data-testid="callout-only-worksheets"
                      >
                        Sheet chỉ xuất hiện ở 1 bên: {workbookOnly.join(", ")}.
                        Lưới bên dưới so sánh cặp sheet bạn đã chọn theo vị trí
                        ô.
                      </div>
                    )}
                    {otherSideSheetMissing && (
                      <div
                        className="only-sheet-callout"
                        data-testid="callout-missing-sheet"
                      >
                        Sheet được chọn bị thiếu ở file còn lại. Tất cả các ô
                        của sheet sẽ hiển thị dạng Thêm mới hoặc Bị xóa.
                      </div>
                    )}
                    <div
                      className="sheet-tabs"
                      aria-label="Sheet đang được so sánh"
                    >
                      <div
                        className="sheet-tab active"
                        data-testid="tab-before-sheet"
                      >
                        <span className="tab-indicator" /> File cũ ·{" "}
                        {result.beforeSheet}
                      </div>
                      <div
                        className={`sheet-tab ${
                          otherSideSheetMissing ? "old-only" : ""
                        }`}
                        data-testid="tab-after-sheet"
                      >
                        <span className="tab-indicator" /> File mới ·{" "}
                        {result.afterSheet}
                      </div>
                    </div>
                    <div className="paired-grids">
                      <div className="grid-half">
                        <div className="grid-side-label">
                          FILE CŨ · {result.beforeSheet}
                        </div>
                        {grid("before", beforeWorksheet)}
                      </div>
                      <div className="grid-half">
                        <div className="grid-side-label">
                          FILE MỚI · {result.afterSheet}
                        </div>
                        {grid("after", afterWorksheet)}
                      </div>
                    </div>
                    <div className="grid-caption">
                      {gridRows > 0 && gridColumns > 0
                        ? `Hiển thị dòng ${gridStartRow + 1}–${
                            gridStartRow + gridRows
                          } và cột ${gridCols[0]}–${
                            gridCols[gridCols.length - 1]
                          }${
                            fullGridRows > gridRows ||
                            fullGridColumns > gridColumns
                              ? " · Xem trước giới hạn 35 dòng và 12 cột quanh ô được chọn"
                              : ""
                          }. Chọn một ô ở danh sách khác biệt bên trái để xem vị trí của nó.`
                        : "Không có dữ liệu trong các sheet này."}
                    </div>
                  </section>
                </div>
              </>
            ) : (
              <div
                className="no-diff-state"
                data-testid="empty-identical-workbooks"
              >
                <span className="no-diff-icon">
                  <CheckCircle2 size={21} />
                </span>
                <strong>Hai sheet này hoàn toàn khớp nhau.</strong>
                <p>
                  Mọi ô dữ liệu đều có giá trị giống nhau. Bạn có thể chọn cặp
                  sheet khác ở trên để tiếp tục kiểm tra.
                </p>
              </div>
            )}
          </section>
        )}
        <div className="footer-note">
          <LockKeyhole size={12} /> File Excel được xử lý trực tiếp trên trình
          duyệt của bạn và sẽ tự động xóa khi bạn bấm Làm mới.
        </div>
      </main>
    </div>
  );
}

export default App;
