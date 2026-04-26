        value_sheet = value_book[sheet_name]
        print_ranges = _normalize_print_ranges(source_sheet)

        if not print_ranges:
            report.skipped_sheets.append(
                SheetResult(name=sheet_name, status="skipped", message="No print area defined.")
            )
            continue

        new_sheet = cleaned_book.create_sheet(title=sheet_name)
        report.processed_sheets.append(_copy_print_area(source_sheet, value_sheet, new_sheet, print_ranges))

    if not report.processed_sheets:
        cleaned_book.create_sheet(title="Summary")
        summary_sheet = cleaned_book["Summary"]
        summary_sheet["A1"] = "No sheets were cleaned because no print areas were defined."
        summary_sheet["A1"].fill = PatternFill(fill_type="solid", fgColor="FFF4CC")

    output_buffer = BytesIO()
    cleaned_book.save(output_buffer)
    output_buffer.seek(0)

    return MemoryCleanResult(
        output_bytes=output_buffer.getvalue(),
        output_filename=output_filename,
        report=report,
        elapsed_seconds=time.perf_counter() - start_time,
        source_size_bytes=len(uploaded_bytes),
        source_sheet_count=len(formula_book.sheetnames),
    )


def _validate_upload(uploaded_file: object | None) -> str | None:
    if uploaded_file is None:
        return None

    filename = getattr(uploaded_file, "name", "")
    if Path(filename).suffix.lower() != ".xlsx":
        return "仅支持 .xlsx 格式文件，请勿上传 .xls、.csv 或其他格式。"

    size = getattr(uploaded_file, "size", None)
    if size is not None and size > MAX_UPLOAD_SIZE_BYTES:
        return "文件超过 100MB 限制"

    return None


def _build_output_filename(filename: str) -> str:
    source = Path(filename)
    stem = source.stem or "cleaned_workbook"
    return f"{stem}_cleaned.xlsx"


def _render_receipt(result: MemoryCleanResult) -> None:
    report = result.report
    receipt_rows = [
        ("处理耗时", f"{result.elapsed_seconds:.2f} 秒"),
        ("清理前文件大小", _format_size(result.source_size_bytes)),
        ("清理后文件大小", _format_size(len(result.output_bytes))),
        ("原始 Sheet 数", str(result.source_sheet_count)),
        ("DNP 保留 Sheet 数", str(len(report.retained_sheet_names))),
        ("DNP 删除 Sheet 数", str(len(report.removed_by_dnp_names))),
        ("实际处理 Sheet 数", str(report.processed_count)),
        ("跳过无打印区域 Sheet 数", str(report.skipped_count)),
        ("清理隐藏行数", str(report.hidden_rows_removed_count)),
        ("清理隐藏列数", str(report.hidden_columns_removed_count)),
        ("公式转空值数", str(report.formula_to_empty_count)),
        ("保留图片数", str(report.images_kept_count)),
    ]

    st.subheader("处理回执")
    metric_cols = st.columns(4)
    metric_cols[0].metric("处理耗时", f"{result.elapsed_seconds:.2f}s")
    metric_cols[1].metric("处理 Sheet", report.processed_count)
    metric_cols[2].metric("跳过 Sheet", report.skipped_count)
    metric_cols[3].metric("隐藏行列", report.hidden_rows_removed_count + report.hidden_columns_removed_count)

    st.table([{"项目": label, "结果": value} for label, value in receipt_rows])

    if report.skipped_sheets:
        skipped_names = "、".join(sheet.name for sheet in report.skipped_sheets)
        st.warning(f"以下 Sheet 因未设置打印区域被跳过：{skipped_names}")


def _format_size(size_bytes: int) -> str:
    units = ("B", "KB", "MB", "GB")
    value = float(size_bytes)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{size_bytes} B"


def _inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: #eef4f7;
        }
        .block-container {
            max-width: 980px;
            padding-top: 3rem;
            padding-bottom: 4rem;
        }
        .auth-shell {
            max-width: 520px;
            margin: 12vh auto 0;
            padding: 2rem;
            border: 1px solid #dce8ee;
            border-radius: 8px;
            background: #ffffff;
        }
        h1, h2, h3 {
            color: #102631;
            letter-spacing: 0;
        }
        div[data-testid="stFileUploader"] section {
            border-color: #b8cbd4;
            background: #ffffff;
        }
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #dce8ee;
            border-radius: 8px;
            padding: 1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
