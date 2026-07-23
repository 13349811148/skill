---
name: data-report-multi-platform
description: "Generate Chinese operating reports for Tmall and Pinduoduo from the local desktop operations database. Use when the user asks for 日报V2版本（天猫+拼多多）, 日报V2, 数据报表多平台版, 天猫拼多多合并报表, 多平台店铺报表, or wants one Skill to generate either a 天猫店铺数据报表 or the previous 拼多多旧格式数据报表 for specified dates."
---

# 日报V2版本（天猫+拼多多）

Use this Skill to generate separate reports for recognized platforms. Before calculating, inspect every workbook sheet, including hidden sheets, and look for a valid order header in its first 20 rows. Recognize 天猫 from its order-date fields together with signals such as `宝贝ID`、`交易状态`、`子订单编号` or `主订单编号`; recognize 拼多多 from `样式ID`, an order date, and a distinct field such as `商家实收金额(元)`、`汇总_店铺名称`、`多多支付立减金额(元)` or `团ID`. Run the matching platform mode automatically. If both platforms are recognized, generate one report for each platform; never combine their orders into one workbook calculation. If no platform is recognized, report the unmatched files and ask for a valid export.

## Update Gate

Before any business workflow, run the bundled `scripts/update-skills.ps1` exactly once for the current user request. Let the script infer the AI client home from this Skill's installed path. If the update succeeds or reports that the installed release is current, reread this `SKILL.md` from disk before continuing so the newest rules take effect. Do not run the update gate again after rereading during the same request. If the update check fails, keep the installed version, tell the user briefly, and continue with the current rules.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<this-skill-directory>\scripts\update-skills.ps1"
```

## Mandatory Execution Order

Follow this order exactly. Do not skip a step, run a platform builder directly, or report completion before the final check.

1. Run the **Update Gate** and reread this file when the installed release is current or updated.
2. Start only through `scripts/build_report.py`; inspect order headers and select the platform automatically, or pass its `--platform` option only when the user explicitly requests one platform.
3. Read and validate the marketing-activity and cost workbooks before calculation. Stop and report the specific issue if a required marketing workbook cannot be read or lacks valid matching records.
4. Let the dispatcher run the internal platform builder. Do not invoke `build_tmall_report.py` or `build_pdd_report.py` as a separate user workflow.
5. Verify that every requested platform report was created, required output fields have been checked, and all warnings have been written to the generation log and sent in the current Codex or Work Buddy conversation.

## Platform Routing

- **天猫**: Run `scripts/build_report.py --platform tmall`. Output `天猫店铺数据报表.xlsx`. When the order export has no `商家实收` field, use `买家应付` as `商家实收金额(元)` for receipt matching, sales statistics, small-receipt checks, and promotion allocation. Keep the current 天猫 rules: common Tmall export headers, and combine 万相台 `花费`、新客加速 `预估新客加速费用`、老客加速 `预估老客加速费用` (plus any legacy 品牌新享费用, which is included under 新客加速) into `推广费用`. In `备注`, show exactly three component amounts: `万相台费用`、`新客加速费用`、`老客加速费用`; 新客加速与老客加速均属于原品牌新享口径. Retain the zero-receipt registration-price adjustment for subsidy activities. Keep every promotion amount keyed by `日期 + 商品ID`, then preserve the existing rule of splitting that product’s fee across its matched 样式ID order rows by actual receipt amount; never allocate a grand total directly or evenly. Match costs using the order's specification-level merchant code and carry the matched cost table's `产线`、`项目组`、`管理类型`、`品种` into the output.
- **拼多多**: Run `scripts/build_report.py --platform pdd`. Output `数据报表拼多多.xlsx`. Keep the prior 拼多多 old-format rules, including style-ID aggregation, marketing/activity fallback, cost lookup, promotion allocation, subsidy, cost, net-profit, and pricing reminders. Always retain the matched marketing-activity row's `活动价`、`报名价` and `促销机制` in every normal sales or empty-burn output row, regardless of whether the row qualifies as an official subsidy activity. Populate `促销机制` from `活动名称`; only fall back to legacy promotion-mechanism fields when `活动名称` is blank. Continue to use the separate subsidy-qualification rule only for subsidy calculations. When the matched marketing row has `单件预估实收金额`, use it as the expected arrival price: if it matches order actual receipt ÷ quantity, fill the expected value; if actual is lower by no more than 5%, fill actual and warn; if actual is lower by more than 5%, leave the arrival-price cell blank and warn; if actual is higher, fill actual and warn that the activity or promotion may have expired. If the column is absent or the matched cell is blank, fill `到手价` with the calculated actual value and warn the user that the marketing activity table lacks `单件预估实收金额`. Record every exception in the generation log and also send the reminder in the current chat window, whether the user is using Codex or Work Buddy; never rely on the log alone.

For both platforms, never transfer a product's promotion spend to another selling product. When promotion data exists for a `日期 + 商品ID` but that product has no effective sales amount for the date, add one independent `空烧推广费` row for that product and date. Keep the promotion fee and every product, shop, category, cost, and activity field that can be matched safely; write both `实际成交数量（去退款去补单后）` and `实际成交金额（去退款去补单后）` as `0`, write net sales and gross profit as `0`, and write net profit as the negative promotion fee. If a product has multiple specification records and the promotion export does not identify one uniquely, leave both `SKU ID` and `商品SKU` blank instead of assigning the spend to an arbitrary specification. Mark `备注` as `空烧推广费：无有效销售`, retain the three Tmall promotion-component amounts in the Tmall remark, and record the generated empty-burn row in the generation log. Keep every normal sales row before every empty-burn row; place the complete empty-burn group at the absolute bottom of the output table, then sort that bottom group by `日期`、`店铺名称`、`商品ID`、`SKU ID`、`商品SKU`.

## 数据整理至日报交接保护

Apply these rules to files produced by the data-organizing Skill and to equivalent original exports:

1. Resolve each valid order row's business date independently in this priority: `订单支付时间` → `支付时间` → `订单付款时间` → `付款时间` → `订单成交时间`; for 天猫 only, finally fall back to `订单创建时间`. Use the first valid value on that row. If all available candidates are blank or invalid, stop and show the file, sheet, physical row, order number, and raw date values; never silently skip it or use the file name as the order date.
2. For a workbook produced by the data-organizing Skill, classify each promotion row first by `归档推广类型`, not by the combined workbook's header. Accept `拼多多-商品推广` only in the Pinduoduo builder; accept `万相台`、`新客加速` and `老客加速` only in the Tmall builder. Preserve each type's original fields when selecting its amount: 万相台 uses `花费`, 新客加速 uses `预估新客加速费用` or legacy 品牌新享 `预估抽佣金额`, and 老客加速 uses `预估老客加速费用`. For an original, unorganized export with no `归档推广类型`, detect platform and type by workbook headers as before. Skip a clearly recognized other-platform row; stop when a promotion-cost row lacks a valid type. Recognize legacy 品牌新享 from `业务日期` together with `预估抽佣金额`, and include it under the new-customer acceleration component.
3. Accept organizer keys `归档日期` and `归档商品ID` first; otherwise accept promotion date aliases `日期`、`业务日期`、`统计日期`、`统计时间`、`推广日期`、`报表日期` and product aliases `商品ID`、`商品id`、`商品Id`、`宝贝ID`、`主体ID`、`商品编号`、`商品编码`、`商品ID（必填）`.
4. Scan all workbook sheets, including hidden sheets, and locate eligible headers within the first 20 rows. Ignore instruction and auxiliary sheets that do not contain a valid required header. Exclude generated order summaries, organizing logs, missing-data reminders, interruption logs, backups, temporary files, and Excel lock files from order discovery.
5. Output a text-formatted `SKU ID` column immediately after `商品ID`. Map 拼多多 `样式ID` and 天猫 `SKU ID`、`SKU_ID`、`SKUID` to it. Keep the existing output header `商品SKU`, but its value must be the specification-level merchant code such as `商家编码-规格维度`、`规格编码` or `SKU编码`; never put a specification name, product name, or title there. Keep different SKU IDs of the same product in separate rows.
6. When an order lacks a specification code, first use 商品数据's exact `商品ID + SKU ID → 规格编码` mapping, then the exact marketing-activity style record. Use a product-level mapping only when it resolves to exactly one specification. For an empty-burn row, fill `SKU ID` and `商品SKU` only when the product has exactly one resolvable specification; otherwise leave both blank. Match the cost table using the final output `商品SKU` code.
7. Before any sales, quantity, order-count, customer-price, product-ranking, promotion-allocation, or other operating calculation, classify every **order detail**. Exclude a detail when `订单性质=补差价` or `是否计销售=否`. For unmarked historical data, standardize `商品标题`/`商品名称` by trimming whitespace, normalizing full-width/half-width characters, and removing square-bracketed marketing prompts; exclude the detail when the standardized title contains `补差价` or `补收差价`. Never infer this from 商品ID、金额、数量 or SKU/specification. A mixed order must retain its ordinary product details. Do not write a separate 补差价 row in the daily workbook; record the excluded detail count and `已识别为补差价，不计销售` in the generation log and current Codex or Work Buddy conversation.

## Promotion Matching Protection

Apply all three safeguards before writing either platform workbook. A failure in any safeguard must stop report generation; never silently skip an invalid nonzero promotion row or continue with an unreconciled amount.

1. **Daily date integrity**: Use a row-level `日期`、`业务日期`、`统计日期`、`统计时间`、`推广日期` or `报表日期` value when the promotion export provides one. Every nonzero-spend row must resolve to exactly one calendar day. Do not require one row for every calendar day in the requested range: a day with no promotion spend may be completely absent. For example, a 1–3 day export containing dated spend rows only for day 1 and day 3 is valid when day 2 had no spend. Never assign a multi-day export total to the range's first date. If the filename or date cell covers multiple days and the file has no single-day detail for every nonzero-spend row, stop and ask for a daily-detail export. This rule applies especially to Pinduoduo recent-seven-day promotion exports.
2. **Promotion snapshot deduplication**: First remove byte-for-byte duplicate files by SHA-256 and keep the newest copy. Then aggregate rows within each source file by `平台 + 店铺 + 推广类型 + 日期 + 商品ID`. Do not require, read, match, infer, or output a plan identity or plan ID; year/month archive folders are never plan identities. If the same complete business key occurs in multiple source files, keep only the newest file by modification time and write the ignored files to the generation log. If missing `店铺` makes overlapping snapshots ambiguous, stop and ask the user to provide the shop dimension or retain only one clearly newest snapshot; never add ambiguous snapshots together.
3. **Input-output reconciliation**: Immediately before workbook writing, reconcile every deduplicated imported promotion amount against the sum of `推广费用` in all normal sales rows plus all `空烧推广费` rows, first by `日期 + 商品ID` and then in total. The maximum permitted absolute difference is `0.01` yuan. If any key or the total exceeds this tolerance, stop, show the imported amount, output amount, and difference, and do not produce the workbook.

## Cost Matching and Data Quality Protection

Apply these rules to both platforms while leaving the existing refund and supplemental-order handling unchanged:

1. Treat `SKU ID` as the platform's specification identity and `商品SKU` as the specification-level merchant code; never treat `商品SKU` as a specification display name. Use only the final `商品SKU` code, sourced from an order field such as `商家编码-规格维度`、`商家编码（规格维度）`、`规格维度商家编码`、`规格商家编码`、`SKU商家编码` or `SKU编码`, for cost-table matching. Match it against a cost-table key column named `商品编码`、`商品SKU名称`、`商品SKU`、`SKU名称` or `SKU编码`. Normalize full-width characters, surrounding or embedded whitespace, and letter case before comparing while preserving the original values in logs.
2. Read every cost-table worksheet up to the first `10,000` rows and `100` columns. Warn when a sheet exceeds this limit. Accept a positive numeric cost from `6.11成本价`、`成本价`、`产品成本` or `成本`; never convert a nonblank invalid value such as `待定` into zero. Stop if the cost workbook cannot be read or contains no usable positive cost.
3. Use costs in this order: matching cost-table record; unique marketing-activity `样式ID`; exact `商品ID + 商品SKU`; same-product fallback; otherwise blank. Carry `项目组`、`管理类型`、`品种`、`产线` from a matched cost-table record for both platforms. For every normal and empty-burn output row, write one structured record to the standalone `<日报文件名>_成本获取日志.csv` with `平台`、`行类型`、`日期`、`店铺名称`、`商品ID`、`样式ID`、`订单规格编码`、`产品成本` and exact `成本来源`. Mark the row type as `正常销售` or `空烧推广费`. Never copy these row-level cost amounts or cost-source records into the generation log or any other log.
4. Treat marketing-activity `样式ID` as unique. If it repeats, continue with the current last-record behavior but prominently warn the user and write the duplicate value and count to the log. Likewise warn when normalized cost-table matching codes repeat.
5. Treat an activity as a subsidy activity only when both prices are valid numbers and `报名价 > 活动价`. If `报名价 < 活动价`, do not calculate a negative subsidy; warn that the marketing data is probably wrong and identify the affected style, product, SKU, and prices. Do not determine subsidy status from promotion-mechanism text.
6. Deduplicate byte-for-byte identical order files, and when the same identified order line occurs across overlapping exports, keep the newest file by modification time using `店铺 + 订单号 + SKU ID + 商品ID + 商品SKU`. Do not merge unrelated rows merely because an order-number field is absent. Stop on an unreadable relevant order or promotion file, a required-field failure, or an invalid nonblank order/promotion number. Keep 商品ID、SKU ID and 商品SKU as text when writing Excel so long identifiers never lose digits.
7. For 商品数据 exports above `100,000` rows, use the first `100,000` rows and warn; do not discard the entire file.

For both platform outputs, display `实际成交数量（去退款去补单后）` as an integer, `每单补贴金额` and `总补贴金额` as numbers with two decimal places, and `毛利率` as a percentage.

For both platform outputs, place the four cost-metadata columns at the absolute end of the workbook after all normal, blank, and auxiliary columns. Keep their exact final order as `项目组`、`管理类型`、`品种`、`产线`; never insert another output column after `产线`.

Both modes use `%USERPROFILE%\Desktop\运营数据库`. They auto-detect the marketing activity and cost workbooks from `营销活动监控`, preferring the most recently modified workbook whose name contains `营销活动_更新` when present; only calculate dates explicitly requested by the user.

Read the marketing activity workbook before calculating. If it cannot be opened, has no `商品ID`、`样式ID`、`商品SKU` headers, or has no valid product records, stop immediately and report the specific failure. Never generate a report with blank marketing fields after a marketing-workbook read failure.

For platform detection and 商品数据 exports, inspect at most the first 100,000 rows. If a 商品数据 export exceeds this limit, report that only the first 100,000 rows were used; do not silently claim that the report covers the entire file. Read all eligible rows from order and promotion exports for the calculation.

## Workflow

Run the dispatcher without `--platform` to inspect the order files automatically. When both Tmall and Pinduoduo tables are present, it generates the two platform reports separately. Use `--platform tmall` or `--platform pdd` only to force one platform.

For one day:

```powershell
$env:PYTHONUTF8='1'
python ".\scripts\build_report.py" --date 2026-07-14
```

For a date range, append `--start-date YYYY-MM-DD --end-date YYYY-MM-DD`. Review the generated warning CSV log and the separate cost-acquisition CSV log beside the selected platform’s output file.

## Order Anomaly Confirmation

Before producing either platform report, scan every eligible order and require confirmation when either condition is true: the order's total actual receipt is greater than `0` and less than `1` yuan; or its actual unit receipt (`订单实收 ÷ 数量`) is lower than the matched unit product cost by more than `10%`. Only review prices below cost; do not stop for prices above cost. Prefer the cost-table match by specification-level merchant/SKU code, then fall back to the marketing-activity product cost. If no positive cost can be matched, do not run the deviation comparison for that order; write a missing-cost warning to the generation log and send it in the current conversation. For a Tmall zero-receipt subsidy order that qualifies for the registration-price adjustment, use the adjusted receipt for the cost-deviation comparison.

If any review condition is found, stop without generating a report, show each affected order and every reason that it matched, and ask the user whether to count those orders.

- After the user confirms inclusion, rerun with `--small-receipt-action include`.
- After the user says not to count them, rerun with `--small-receipt-action exclude`; exclude those orders from both quantity and sales amount.
- If one order matches both conditions, show it once with both reasons and apply the user's decision to the whole order.
- Never silently include or exclude reviewed orders. The default `confirm` action must remain in effect until the user answers.

## Author Update Publishing

Apply this section only when a user asks to modify or update this Skill itself; never run it during ordinary report generation. Allow any user to edit and validate the installed local Skill. Before any cloud release mutation, use the repository-root cross-platform `publish-skills.py`; its mandatory author gate must verify all three conditions: the authenticated GitHub login is exactly `13349811148`, `origin` is exactly `13349811148/skill`, and the login has `ADMIN` permission on that repository. Never bypass, weaken, or replace this gate with Git commit name or email checks.

When the author gate succeeds, treat the author's request to modify or update this Skill as authorization to finish the complete release after validation. The word `推` means commit and push. Let the publisher copy the installed Skill pair, generate the next release version, commit the intended changes, and push `main`; then confirm the remote commit and clean final status.

When the author gate cannot verify all three conditions or returns `PUBLISH_BLOCKED_LOCAL_ONLY`, preserve the user's installed local changes and do not create a release version, repository commit, or push. In the current Codex or Work Buddy window, prominently tell the user: `当前修改仅对这台电脑上的本地 Skill 生效。云端版本未更新；如需云端更新，请联系作者。` Then use AI to polish the actual local rule changes into one self-contained Chinese prompt inside a copyable fenced code block and send it to the non-author user. Address the prompt to the Skill author; include the Skill name, affected platform or workflow, current problem, exact requested rules and edge cases, expected acceptance checks, and a request to validate and publish the cloud version. Base it on the real local changes or the user's request when no baseline diff exists. Exclude credentials, tokens, private store data, and machine-specific paths. Do not contact the author automatically; the non-author user decides whether and how to forward the prompt.

## Resources

- `scripts/build_report.py`: Platform dispatcher.
- `scripts/build_tmall_report.py`: 天猫 calculation rules.
- `scripts/build_pdd_report.py`: 拼多多旧格式 calculation rules.
- `scripts/data_quality_protection.py`: Shared strict-number, cost-code normalization, duplicate-warning, price-validation, and standalone cost-acquisition log rules.
- `scripts/handoff_protection.py`: Shared order-date priority, cross-platform header classification, multi-sheet header search, file exclusion, and product/SKU mapping rules.
- `scripts/promotion_protection.py`: Shared daily-date integrity, promotion-snapshot deduplication, and input-output reconciliation safeguards.
- `scripts/update-skills.ps1`: Check the shared public GitHub repository and transactionally update both Skills before business work.
- `references/template_columns.csv`: Shared output-column order.
