---
name: shop-report-organizer
description: "固定口令触发：仅当用户明确输入完整短语“整理数据库V2版本”时使用；没有这个完整短语时不要使用。触发后按本 Skill 的流程执行。"
---

# 整理数据库V2版本

Use this skill to organize files downloaded from shop backends into the operations database. The workflow supports multiple platforms, currently Pinduoduo and Tmall. It must inspect spreadsheet content first; do not identify report type from filename keywords alone.

## Update Gate

Before any business workflow, run the bundled `scripts/update-skills.ps1` exactly once for the current user request. Let the script infer the AI client home from this Skill's installed path. If the update succeeds or reports that the installed release is current, reread this `SKILL.md` from disk before continuing so the newest rules take effect. Do not run the update gate again after rereading during the same request. If the update check fails, keep the installed version, tell the user briefly, and continue with the current rules.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<this-skill-directory>\scripts\update-skills.ps1"
```

## Mandatory Execution Order

Follow this order exactly. Do not scan, classify, preview, or archive downloaded reports before the ZIP preparation step succeeds.

1. Run the **Update Gate** and reread this file when the installed release is current or updated.
2. Confirm the database root and input folder, then run only `scripts/organize_reports.py --dry-run` as the entry point. Its first input-preparation action must safely extract every ZIP in `下载未分类` and move each successfully processed ZIP to `_已解压ZIP`.
3. If any ZIP cannot be extracted, contains no supported spreadsheet, or fails the safety check, stop immediately, report the ZIP filename and error in the chat, and do not scan the remaining reports or run `--apply`.
4. After ZIP preparation succeeds, let the script inspect workbook contents, classify reports, and compare file contents. For files with identical SHA-256 content, retain one deterministic source for processing and mark every other copy as `duplicate_file`; do not classify reports from filenames alone.
5. Review every preview/reminder row with the user. Run `--apply` only after the preview is acceptable and all required reports are present, unless the user explicitly approves `--allow-missing-required` for a supplemental batch.
6. Verify the apply result, including every `duplicate_file` moved to `下载未分类\_重复文件`, then send every required reminder and unresolved item in the current Codex or Work Buddy conversation.

## Workflow

1. Confirm the source folder and database root. Default to:
   - Database root: `%USERPROFILE%\Desktop\运营数据库`
   - Input folder: `%USERPROFILE%\Desktop\运营数据库\下载未分类`
   - Resolve `%USERPROFILE%` to the current Windows user's profile directory. Do not hard-code another person's Windows username.
   - If a `.zip` file is present in `下载未分类`, the organizer must safely extract contained `.csv`, `.xlsx`, or `.xls` files into `下载未分类` before scanning them. After successful extraction, move the original ZIP into `下载未分类\_已解压ZIP` so it will not be extracted again.
   - ZIP extraction is an input-preparation step and therefore also runs before `--dry-run`. Preview still must not rename, archive, or merge the extracted report files.
2. Run a preview first. The script automatically prepares any ZIP files as described above. When this skill is packaged in a plugin, use the bundled
   `scripts/organize_reports.py` from this skill directory:

```powershell
$databaseRoot = Join-Path $env:USERPROFILE "Desktop\运营数据库"
python ".\scripts\organize_reports.py" --database-root "$databaseRoot" --dry-run
```

3. Review `整理记录_预览.csv`, `缺失提醒_预览.csv`, and the console summary. Pay attention to files marked `pending` or `duplicate_file` and reminders marked `missing`. Copy every row of the reminder log into the chat response as required by **Chat Reminder Output** below.
4. For order reports, remember that each download usually covers the latest 7 days, so repeated downloads overlap. On `--apply`, merge order rows into cumulative shop/month tables by transaction month.
   - Use at most the first 100,000 rows only for preview, report recognition, and shop/period inference. This is a script-side preview guard, not an AI row-processing limit.
   - During every formal cumulative merge or rebuild, read every row and every effective spreadsheet column in every source and existing cumulative table. This applies to order monthly tables, after-sale cumulative tables, promotion-adjustment cumulative logs, and order-summary rebuilds. Never apply the 100,000-row or 120-column preview guards to formal processing, and never load the entire table into the AI context merely for reasoning.
5. Only after the preview is acceptable and no required report is missing, run:

```powershell
$databaseRoot = Join-Path $env:USERPROFILE "Desktop\运营数据库"
python ".\scripts\organize_reports.py" --database-root "$databaseRoot" --apply
```

For supplemental batches that intentionally contain only one report type, use `--allow-missing-required` on `--apply` after confirming the preview.

## Naming And Filing

### Platform And Shop Folder Naming

- Keep existing Pinduoduo shop folder names unchanged.
- For every Tmall report, normalize the shop name to `店铺名称（天猫）` before building target folders or cumulative filenames. Example: `庄芯粮油专营店（天猫）`.
- Detect platform from the matched report signature first, then from explicit `天猫` or `Tmall` text in the filename or spreadsheet. A product-ID mapping or existing archive folder whose shop name already ends in `（天猫）` also preserves the Tmall marker.
- Never guess that an unknown report is Tmall merely because it does not match a Pinduoduo signature. Leave it pending and show the source file in the chat reminder log.
- Product IDs learned from Tmall product data must map to the suffixed shop name so later Tmall order and promotion exports keep using the same folder.

Use this filename format:

```text
下载日期-店铺名称-报表类型-包含时段.原扩展名
```

Special filename format and merge rule for `推广调整日志`:

```text
店铺—商品ID—推广调整日志更新至YYYY-MM-DD.xlsx
```

For `推广调整日志`, keep one cumulative table per shop and product ID. On `--apply`, read every row from the new source and every existing cumulative table without the 100,000-row preview cap, merge them, de-duplicate identical operation rows, sort by `操作时间` newest first, and rename the cumulative table so the filename's `更新至YYYY-MM-DD` matches the latest operation date in the merged rows. Print the merged target file, source file count, and de-duplicated row count in the message window.

Special filename format and merge rule for `售后数据`:

```text
店铺—售后数据更新至YYYY-MM-DD.xlsx
```

For `售后数据`, keep one cumulative table per shop in the current download-month archive folder. On `--apply`, read every row from the new source and every existing cumulative table without the 100,000-row preview cap, merge them, de-duplicate by `售后编号`, keep the row from the later source when the same `售后编号` appears again, sort by `申请时间` newest first, and rename the cumulative table so the filename's `更新至YYYY-MM-DD` matches the latest `申请时间` in the merged rows. Print the merged target file, source file count, and de-duplicated row count in the message window.

Special validation and merge rule for `推广数据`:

- Require every non-summary promotion detail row to contain both a valid specific date and a valid product ID. Never use a filename date as a fallback for promotion archiving. If either field is missing from the header or any detail row, keep the source in its original location, mark it `pending`, and tell the user to re-download the promotion detail with both fields.
- Keep exactly one cumulative promotion workbook per shop at `推广数据/<店铺>/<店铺>—推广数据.xlsx`; do not create one workbook per product ID.
- Preserve each source column and add normalized `归档日期`、`归档商品ID` and `归档推广类型` columns, so the combined shop workbook always contains the date, product ID and promotion type used for filing. Write Pinduoduo rows as `拼多多-商品推广`; write Tmall rows as `万相台`、`新客加速` or `老客加速`. Classify legacy 品牌新享 as `新客加速` while preserving its original columns.
- On `--apply`, validate every row of the new source before moving any file. Then merge the source with the shop's existing cumulative promotion workbook, de-duplicate only completely identical rows, sort by `归档日期`、`归档商品ID` and `归档推广类型`, and replace the workbook transactionally. Print the target file, source file count, and de-duplicated row count in the message window.

Use the existing archive layout:

```text
报表大类/店铺/年份/月（导出时间）/文件
```

If any required field cannot be determined, put the file under `_待确认` and explain the missing field in the manifest. Do not guess silently.

## Missing Report Reminders

Every normal preview or apply run must check `references/expected_reports.csv` after reading the input folder.

Rules:

- Treat rows with `required=是` as required report types.
- If a required report type is not detected in the current input batch, print a reminder message and write it to:
  - `缺失提醒_预览.csv` during `--dry-run`
  - `缺失提醒.csv` during `--apply`
- During `--apply`, if any required report type is missing, stop before moving files. Write `整理记录_中断.csv` and `缺失提醒.csv`, print the reminders, and do not rename, move, archive, update `shop_product_map.csv`, or update order summaries.
- During `--dry-run`, missing required reports are reminders only because preview mode never moves files.
- Count a report type as present when it is recognized from table content, even if the file later becomes `pending` because store name or period is missing.
- Do not require `商品数据` by default, because product files are not exported every day or every month.
- If `订单数据` or `推广数据` contains a product ID that is not found in `references/shop_product_map.csv` or already archived `商品数据/<店铺>/...`, print a reminder to update or re-export that shop's `商品数据`.
- If only `售后数据` contains a product ID that is not found in the product map or archived product data, treat it as a historical or likely delisted product. Write it as `status=historical_product_id`; do not ask the user to re-export product data for that reason alone.
- Write new-product and historical-product reminders into the same reminder files, including `shop_name`, `source_name`, and `new_product_ids`. These reminders must also be printed in the message window every run, so the user does not need to open the CSV to understand what is missing or only historical.
- Do not run missing-report reminders during `--rebuild-order-summary`, because that mode rebuilds summaries from archived files and does not inspect a new input batch.
- To change which report types trigger reminders, edit `references/expected_reports.csv`.

## Chat Reminder Output

The CSV reminder files are an audit copy, not the primary way the user should learn the result. After every `--dry-run` and `--apply` run:

- Show every reminder-log row in the chat window. This includes `missing`, `new_product_id`, `historical_product_id`, `duplicate_file`, `pending`, and `ok` rows.
- For each row, include its status, shop name, report type, source filename, product IDs when present, and the complete `message` text. Do not replace these rows with only a count, a CSV path, or a general summary.
- When the only row is `status=ok`, explicitly tell the user that no required report, new-product, historical-product, or pending-file reminder was found.
- Show the complete reminder log after preview. If apply is run, show the complete apply reminder log in the final response; do not assume the preview message was enough.
- If console output is truncated or has encoding problems, read `缺失提醒_预览.csv` or `缺失提醒.csv` as UTF-8 and relay every row from the file.
- Keep the response readable, but never require the user to open a spreadsheet to discover any reminder detail.

## Recognition Rules

Read `references/report_signatures.csv` for report signatures. Current content-based signatures:

- Pinduoduo `订单数据`: headers include `订单号`, `订单状态`, `商品id`, plus either `支付时间`/`订单支付时间` or `订单成交时间`.
- Tmall `订单数据`: headers include `子订单编号`, `主订单编号`, `订单状态`, `商品ID`, plus either `订单付款时间`/`订单支付时间` or `订单成交时间`.
- Pinduoduo `售后数据`: headers include `售后编号`, `退款类型`, `申请时间`, `商品ID`.
- Tmall 退款 `售后数据`: headers include `订单编号`, `退款编号`, `退款申请时间`, `退款状态`, `商品id`; use the refund application time as the period and the refund number to de-duplicate cumulative records.
- `商品数据`: headers include `商品ID（必填）` and `商品名称`. Current product exports may be either the old product-code table with `商品编码`, or the newer stock export table with `SKUID（必填，注意不是SKU编码）`, `规格名称`, `库存增减`, and `规格编码`.
- Tmall 商品发布模板 is also `商品数据`: recognize the headers `商品Id`, `类目id`, `商品标题`, `一口价`, `skuId`, `价格（元）`, and `数量`, including when the usable table is in a hidden or protected worksheet; infer the shop from a filename parenthesis such as `（天猫-店铺名）`.
- Pinduoduo `推广数据`: headers include `商品名称`, `商品ID`, `推广名称`, `总花费(元)`, `净实际投产比`.
- Tmall 万相台 `推广数据`: headers include `日期`, `主体ID`, `主体类型`, `主体名称`, `花费`, `投入产出比`.
- Tmall 品牌新享 `推广数据`: headers include `商品名称`, `商品ID`, `业务日期`, `预估抽佣金额`, `支付金额`.
- Tmall 老客加速订单 `推广数据`: headers include `商品名称`, `父订单ID`, `商品ID`, `支付金额`, `预估老客加速费用`, `日期`.
- Tmall 新客加速订单 `推广数据`: headers include `商品名称`, `父订单ID`, `商品ID`, `支付金额`, `预估新客加速费用`, `日期`.

Store-name detection order:

1. Explicit shop text in spreadsheet cells, canonicalized against known shop folders and `references/shop_product_map.csv`.
2. Product ID majority vote from `references/shop_product_map.csv`.
3. Product ID majority vote from already archived files under `商品数据/<店铺>/...`; this lets future runs work even when `商品数据` is not exported again.
4. For `商品数据` only, use the shop name inside filename parentheses to seed or update the product map. A `天猫-` prefix inside the parentheses is a platform label, not part of the shop name.
5. If still unknown, mark the file `pending`.

Period detection:

- For every order row, prefer a valid `订单支付时间`、`支付时间` or `订单付款时间`; when that row's preferred value is blank or invalid, fall back to `订单成交时间`; for Tmall, fall back once more to `订单创建时间`. If both payment and transaction times are valid, use the payment/付款 time. A valid fallback means the row is valid and must not stop processing.
- Pinduoduo `售后数据`: min and max of `申请时间`; Tmall refund exports use `退款申请时间`, falling back to `退款完结时间`.
- `推广数据`: use only the actual spreadsheet date column. Do not use a filename date range as a fallback.
- `推广数据`: archive single-day or multi-day detail only when every non-summary row has a valid date and product ID. Keep a file pending when its date or product ID field is missing, blank, or invalid in any detail row, and tell the user to re-download a promotion detail export that includes both fields. The filename (including whether it contains `分天数据`) is not a condition.
- `商品数据`: use the download date as a one-day period when no content period exists.

Download date comes from the source file's filesystem creation time: use `st_birthtime` on macOS and the Windows creation time on Windows; fall back to modified time only when the platform has no creation-time field. Dates written in filenames never replace the filesystem creation time.

## 补差价订单明细保护

Apply this rule to every current and future platform's order-detail import. Classify one order detail at a time; never classify, exclude, or retain an entire order merely because another detail has the same order number.

- Standardize `商品标题`/`商品名称` (and equivalent title fields) by trimming surrounding whitespace, normalizing full-width and half-width characters, and removing square-bracketed marketing prompts. When the standardized title contains `补差价` or `补收差价`, mark that detail as `订单性质=补差价` and `是否计销售=否`. Do not use 商品ID、金额、数量 or SKU/specification for this decision.
- Preserve the marked detail in its raw monthly order archive. Set ordinary details to `订单性质=正常` and `是否计销售=是` unless the source already provides a non-sales mark.
- Exclude only the marked detail from new-product-ID reminders and product-map requirements. Keep product IDs available for shop inference; do not create a new-product reminder, require a product export, or add a product mapping solely for a 补差价 detail.
- In preview manifests, reminder logs, formal merge output, and the current Codex or Work Buddy conversation, state `已识别为补差价，不计销售`, include the detail count, and state that normal details in a mixed order remain unchanged.

## Order Data Monthly Merge

Maintain de-duplicated monthly order tables by shop and transaction month:

```text
订单数据/店铺/年份/月（成交时间）/店铺—订单数据YYYY年MM月.xlsx
```

Rules:

- Only update monthly order tables during `--apply`; `--dry-run` must not change them.
- Split each order row by the first valid time under the payment/付款 → transaction → Tmall creation fallback rule above. Never split by file creation time, download time, or export month.
- Treat 100,000 rows only as the preview/classification guard. Stream every row during formal order merging, including all rows from an existing monthly table, so 200,000- or 300,000-row inputs are processed by local Python rather than by the AI context.
- Inspect every worksheet during formal order processing. Process each worksheet that independently contains a valid order ID and order-time header, preserve its own header mapping, and ignore instruction or auxiliary worksheets that are not order tables. Never concatenate a second worksheet's header or notes into the first worksheet's order rows.
- Ignore completely blank rows and explicit `总计`、`合计`、`汇总`、`注` or `说明` rows. Also ignore an order only when it has no valid order time and its `订单状态`/`交易状态` explicitly shows `已取消`、`待付款`、`待支付`、`未付款`、`交易关闭` or `已关闭`; do not include it in the cumulative order table. If a real order row still has no valid value in any available order-time field, stop the entire apply before moving files, show the source file, row number, order number, and original time values, and keep the source and existing monthly tables unchanged.
- Within one source file, silently de-duplicate only rows whose normalized order fields are completely identical. If the same effective order ID appears again with any different field, stop before moving files and show the worksheet, both row numbers, and the differing field values for user confirmation; never let physical row order decide which conflicting record wins.
- Before deleting a source, reconcile each input's nonblank row count as `有效订单行 + 明确忽略的说明/合计行`, then reconcile the de-duplicated database row count against all rows written to the monthly output parts. Any difference must stop the apply without deleting the source.
- Keep each `.xlsx` worksheet within Excel's limit of 1,048,576 total rows including the header. When one shop-month exceeds 1,048,575 data rows, split it into sequential files named `..._第01部分.xlsx`, `..._第02部分.xlsx`, and so on, and print every part with its row count.
- Keep one cumulative order table per shop per transaction month. If one downloaded order report contains multiple transaction months, split its rows into the matching monthly tables.
- Resolve the unique key for every real order row, never just once from the file header. For Pinduoduo, `订单号` must be non-empty. For Tmall, use that row's non-empty `子订单编号`; when it is blank, fall back to the same row's `主订单编号`. If the required Pinduoduo ID is blank, or both Tmall IDs are blank, stop the entire apply before moving files, show the source file, row number, and original ID field values, and keep the source and existing monthly tables unchanged. Never invent `_row_...` or any other temporary order ID.
- For Tmall, retain both `子订单编号` and `主订单编号` as identity aliases. When an older row used the main ID because its child ID was blank and a later row supplies exactly one child ID for that main ID, merge the old record into the child ID and keep only one order. A main ID may legitimately contain multiple distinct child IDs; keep those child orders separate. If a row without a child ID could match multiple child orders, or a main-to-child fallback is otherwise ambiguous, stop and show the IDs and source rows for user confirmation.
- Enforce order-ID uniqueness across every transaction month of the same shop, not only within one month. Before applying a new order report, search the shop's existing cumulative monthly tables for IDs contained in the new source. If a later source changes an order's effective month under the payment/付款 → transaction → Tmall creation rule, remove that order from its old monthly table and write the merged record only to the new monthly table. Preserve historical non-empty fields and protected buyer information during the move. Rewrite only the affected old and new months, delete an old monthly output when no orders remain, and stage and reconcile all affected outputs before replacing any existing table or deleting the source.
- Original order reports do not need or contain a download-time column. For each raw order report, read its filesystem creation time and write that value to `汇总_最新下载时间`; never use a date parsed from the filename to decide report freshness.
- When the same order appears in multiple downloaded order reports, compare the source creation time per order and keep ordinary fields from the latest source. When reading an existing cumulative monthly table, reuse each row's saved `汇总_最新下载时间` and `汇总_最新来源文件`; never use the cumulative workbook's own creation time as the row's source time. If two raw sources have the same creation time, the current input wins.
- Treat a legacy cumulative row as having unknown source time when its saved source is empty, names the cumulative workbook itself, or otherwise uses the cumulative monthly filename pattern. Keep that historical row, but let any newly downloaded raw report with the same order ID win ordinary fields and replace the legacy metadata with the raw file's creation time and filename.
- Preserve existing non-empty order information when the latest source has no value. If the latest source does not contain a column at all, or contains the column but the order's cell is empty, retain that column's non-empty value from the older record. Use the latest value only when it is non-empty. Apply the protected buyer-information rule below after these missing-column and empty-cell checks.
- Protect buyer receiving information fields from encrypted later reports. For fields such as `消费者资料`, `省`, `市`, `区`, `用户购买手机号`, `收件/收货地址`, `手机号`, and `电话`, if the latest report value is empty or contains `*`, `＊`, or `加密`, keep the old unmasked value in the monthly table. This avoids losing address/contact details after an order changes from `待发货` or `已发货` to `已收货` or `已收货已退款`.
- Keep all source order-report columns, and append metadata columns:
  - `汇总_店铺名称`
  - `汇总_最新下载时间`
  - `汇总_最新来源文件`
  - `汇总_最新归档文件`
- If a future order report has new columns, append those columns to the monthly table.
- Name the worksheet inside every generated order monthly workbook `订单数据`. Keep the real `推广调整日志` report and its worksheet name unchanged.
- Treat source removal and all affected monthly-table replacements as one transaction. Move the source and old outputs into the transaction backup before installing staged outputs; if any move or install fails, restore the source and every old monthly table before reporting failure.
- When running `--rebuild-order-summary`, group and de-duplicate summary rows by each row's payment/付款 → transaction → Tmall creation month, not by the archive file's creation or download month. Preserve row-level source freshness metadata and apply the same cross-month and Tmall alias rules used by the formal monthly tables.
- Print every updated monthly order table in the message window, including the target file, source file count, and de-duplicated row count.

## Resources

- `scripts/organize_reports.py`: preview or apply classification, renaming, movement, manifest generation, and cumulative order/after-sale/promotion-adjustment updates.
- `scripts/update-skills.ps1`: check the shared public GitHub repository and transactionally update both Skills before business work.
- `references/report_signatures.csv`: content signatures and period rules.
- `references/shop_product_map.csv`: product ID to shop name map seeded from the provided examples and updated from newly archived product data.
- `references/expected_reports.csv`: required report list for missing-file reminders.

When adding a new report type, update `report_signatures.csv` first and test with `--dry-run`. When new products appear, add their product IDs to `shop_product_map.csv`, or run an identified `商品数据` file with `--apply` to update the map automatically.
