# Lessons Learned

### [2026-02-19] — Dashboard data completeness
When embedding data for dashboards, always verify the dataset is complete. The original dashboard only had top 200 claims (pre-filtered to loss > 0), which made KPIs like loss ratio (910%) and claims frequency (100%) completely misleading. Always embed the full dataset or proper pre-computed aggregates.

### [2026-02-19] — Compact JSON for embedding
For embedding large datasets in HTML files, use compact array format instead of objects. 10K records as arrays = 294KB vs ~600KB+ as objects with named keys. Use separators=(',', ':') in json.dump.

### [2026-02-19] — matplotlib FormatStrFormatter vs FuncFormatter
`FormatStrFormatter` uses old-style `%` string formatting, which does NOT support the `,` thousands separator (e.g., `'$%,.0f'` will throw `ValueError`). Use `FuncFormatter(lambda x, p: f'${x:,.0f}')` instead for f-string formatting with comma separators.

### [2026-02-19] — PDF page layout budget estimation
When using reportlab platypus with flowables, always estimate total content height per page before building. A US Letter page (612x792 pts) with 0.65in top + 0.7in bottom margins gives ~695 pts usable. Charts at 3.7in = 266 pts each, so two charts plus header + narrative fills a page completely. Keep charts at 3.0-3.3in when sharing a page with tables or long text.

### 2026-02-19 — Session Recovery
When a session is interrupted mid-task, the context summary provides enough detail to continue. Always check the current file state before making edits to avoid duplicating work already done.

### 2026-02-19 — python-pptx Presentation Workflow
On Windows without Node.js, python-pptx is the reliable fallback for PowerPoint generation. Helper functions (add_bg, add_shape, add_text_box, add_multiline) keep slide creation DRY. Use CategoryChartData for bar/column charts.
