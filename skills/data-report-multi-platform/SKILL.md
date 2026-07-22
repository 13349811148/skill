---
name: data-report-multi-platform
description: "Generate Chinese operating reports for Tmall and Pinduoduo from the local desktop operations database. Use when the user asks for 日报V2版本（天猫+拼多多）, 日报V2, 数据报表多平台版, 天猫拼多多合并报表, 多平台店铺报表, or wants one Skill to generate either a 天猫店铺数据报表 or the previous 拼多多旧格式数据报表 for specified dates."
---

# 日报V2版本（天猫+拼多多）

Use this Skill to generate separate reports for recognized platforms. Before calculating, inspect order-export headers: treat `宝贝ID`, `交易状态`, `订单付款时间`, `订单创建时间`, `SKU ID`, or `SKU_ID` as 天猫 signals; treat `样式ID`, `商家实收金额(元)`, or `汇总_店铺名称` as 拼多多 signals. Run the matching platform mode automatically. If both platforms are recognized, generate one report for each platform; never combine their orders into one workbook calculation. If no platform is recognized, report the unmatched files and ask for a valid export.

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

- **天猫**: Run `scripts/build_report.py --platform tmall`. Output `天猫店铺数据报表.xlsx`. When the order export has no `商家实收` field, use `买家应付` as `商家实收金额(元)` for receipt matching, sales statistics, small-receipt checks, and promotion allocation. Keep the current 天猫 rules: common Tmall export headers, and combine 万相台 `花费`、新客加速 `预估新客加速费用`、老客加速 `预估老客加速费用` (plus any legacy 品牌新享费用, which is included under 新客加速) into `推广费用`. In `备注`, show exactly three component amounts: `万相台费用`、`新客加速费用`、`老客加速费用`; 新客加速与老客加速均属于原品牌新享口径. Retain the zero-receipt registration-price adjustment for subsidy activities. Keep every promotion amount keyed by `日期 + 商品ID`, then preserve the existing rule of splitting that product’s fee across its matched 样式ID order rows by actual receipt amount; never allocate a grand total directly or evenly. Match the cost table by order `商家编码`/`SKU编码` against cost-table `商家编码`/`商品编码`, and carry `产线`、`项目组`、`管理类型`、`品种` into the output.
- **拼多多**: Run `scripts/build_report.py --platform pdd`. Output `数据报表拼多多.xlsx`. Keep the prior 拼多多 old-format rules, including style-ID aggregation, marketing/activity fallback, cost lookup, promotion allocation, subsidy, cost, net-profit, and pricing reminders. Always retain the matched marketing-activity row's `活动价`、`报名价` and `促销机制` in every normal sales or empty-burn output row, regardless of whether the row qualifies as an official subsidy activity. Populate `促销机制` from `活动名称`; only fall back to legacy promotion-mechanism fields when `活动名称` is blank. Continue to use the separate subsidy-qualification rule only for subsidy calculations. When the matched marketing row has `单件预估实收金额`, use it as the expected arrival price: if it matches order actual receipt ÷ quantity, fill the expected value; if actual is lower by no more than 5%, fill actual and warn; if actual is lower by more than 5%, leave the arrival-price cell blank and warn; if actual is higher, fill actual and warn that the activity or promotion may have expired. If the column is absent or the matched cell is blank, fill `到手价` with the calculated actual value and warn the user that the marketing activity table lacks `单件预估实收金额`. Record every exception in the generation log and also send the reminder in the current chat window, whether the user is using Codex or Work Buddy; never rely on the log alone.

For both platforms, never transfer a product's promotion spend to another selling product. When promotion data exists for a `日期 + 商品ID` but that product has no effective sales amount for the date, add one independent `空烧推广费` row for that product and date. Keep the promotion fee and every product, shop, category, cost, and activity field that can be matched safely; write both `实际成交数量（去退款去补单后）` and `实际成交金额（去退款去补单后）` as `0`, write net sales and gross profit as `0`, and write net profit as the negative promotion fee. If a product has multiple SKUs and the promotion export does not identify one, leave `商品SKU` blank instead of assigning the spend to an arbitrary SKU. Mark `备注` as `空烧推广费：无有效销售`, retain the three Tmall promotion-component amounts in the Tmall remark, and record the generated empty-burn row in the generation log. Keep every normal sales row before every empty-burn row; place the complete empty-burn group at the absolute bottom of the output table, then sort that bottom group by `日期`、`店铺名称`、`商品ID`、`商品SKU`.

Treat an activity as a subsidy activity only when both `活动价` and `报名价` are present and their values differ. Do not determine subsidy status from the promotion-mechanism text.

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

For a date range, append `--start-date YYYY-MM-DD --end-date YYYY-MM-DD`. Review the generated CSV log beside the selected platform’s output file.

## Order Anomaly Confirmation

Before producing either platform report, scan every eligible order and require confirmation when either condition is true: the order's total actual receipt is greater than `0` and less than `1` yuan; or its actual unit receipt (`订单实收 ÷ 数量`) is lower than the matched unit product cost by more than `10%`. Only review prices below cost; do not stop for prices above cost. Prefer the cost-table match by merchant/SKU code, then fall back to the marketing-activity product cost. If no positive cost can be matched, do not run the deviation comparison for that order; write a missing-cost warning to the generation log and send it in the current conversation. For a Tmall zero-receipt subsidy order that qualifies for the registration-price adjustment, use the adjusted receipt for the cost-deviation comparison.

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
- `scripts/update-skills.ps1`: Check the shared public GitHub repository and transactionally update both Skills before business work.
- `references/template_columns.csv`: Shared output-column order.
