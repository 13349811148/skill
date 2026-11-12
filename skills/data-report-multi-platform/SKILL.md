---
name: data-report-multi-platform
description: "Generate Chinese operating reports for Tmall and Pinduoduo from the local desktop operations database. Use when the user asks for 日报V2版本（天猫+拼多多）, 日报V2, 数据报表多平台版, 天猫拼多多合并报表, 多平台店铺报表, or wants one Skill to generate either a 天猫店铺数据报表 or the previous 拼多多旧格式数据报表 for specified dates."
---

# 日报V2版本（天猫+拼多多）

## 固定报表保存（2026-10-28用户确认）

默认不再按每次请求的日期新建交付文件。商品日报使用固定工作簿，每月新建一个 `YYYY-MM` 子表；店铺日报每月固定一个 XLSX 文件，月内各日累积在“①平台日报”。同一业务日期重跑要替换旧记录，不能重复追加或保留已消失的商品行。其他日期和月份保留。先读 [固定保存与历史迁移](references/report_archive.md)。本节覆盖下文旧的单次输出文件命名说明，计算口径、异常订单确认和推广核对保持原规则。

Use this Skill to generate separate reports for recognized platforms. Before calculating, inspect every workbook sheet, including hidden sheets, and look for a valid order header in its first 20 rows. Recognize 天猫 from its order-date fields together with signals such as `宝贝ID`、`交易状态`、`子订单编号` or `主订单编号`; recognize 拼多多 from `样式ID`, an order date, and a distinct field such as `商家实收金额(元)`、`汇总_店铺名称`、`多多支付立减金额(元)` or `团ID`. Run the matching platform mode automatically. If both platforms are recognized, generate one report for each platform; never combine their orders into one workbook calculation. If no platform is recognized, report the unmatched files and ask for a valid export.

## 商品输出格式（2026-11-07用户确认）

商品日报固定22列，顺序以 [列模板](references/template_columns.csv) 为准。前6列为“日期、店铺名称、商品ID、SKU ID、SKU编码、SKU名称”，SKU名称后增加“货品名称、货品数量”；备注为最后一列。SKU ID只出现一次，紧随商品ID，以文本保存。日期以Excel数值序列保存为真实日期，统一显示`yyyy-mm-dd`，例如`2026-11-07`；保留每行各自的业务日期，不导出ISO时间戳或T00:00:00.000Z后缀。

不输出品牌、所属运营、平台、大类、促销机制，及备注之后原有空白列、大类辅助列、辅助、项目组、管理类型、品种、产线。已删除的活动价、报名价、补贴和净销售额等列也不恢复。原始营销信息和成本元数据继续用于内部匹配、订单校验及独立日志，删除输出列不改变成本、运费、销量、推广、退款和利润口径。

首次更新既有32列月表时，按 [固定保存与历史迁移](references/report_archive.md) 备份后迁移为22列分区表，保留未请求日期、全部月份、业务键和date_notes。只迁移已识别的旧表头；人工公式、对象、批注、重复业务键或未知表头停止，不能重导订单恢复已剔除数据。

## 货品拆分、分区与店铺利润（2026-11-09用户确认）

生成日报默认同步生成固定的`货品日报_拼多多.xlsx`或`货品日报_天猫.xlsx`，每月一个`YYYY-MM`子表。先读[货品与利润规则](references/goods_reporting.md)。商品日报仍一行一个SKU业务键，SKU名称后的货品名称/数量按子件对齐换行；货品日报按日期+店铺+货品编码汇总拆分后的基础单位数量。货品日报增加销售额、毛利、推广费用、净利润、客单价（货品均价）。SKU金额只归属一次：单一货品全额归属；混合装按各子货品产品成本×数量占比分摊，成本不齐时对应混合装金额留空，不能按数量或均分替代。客单价按用户最新确认使用销售额÷货品数量，标题明确标注货品均价。产品/货品表第1行为分区、第2行为字段，第3行起为数据；店铺表分区在第2行，字段仍在第3行。

每次商品计算成功后，由入口调用货品与利润扩展；店铺末尾“店铺净利润”汇总同日同店的SKU净利润，包含空烧推广负利润，其后的“净利率”按下述净成交额口径计算。成本未知则对应日利润汇总留空，不能把缺失当0。用户本次临时确认ZB03、CB10缺成本期间净利润计0；ZB27为ZB03两袋装，同样暂计0，备注必须注明，成本和毛利仍留空，补齐成本后正常计算。

拼多多商品`991147507035`、`1021621380194`按用户最新确认全日期排除销售（不再只限10月27日），推广按实际保留；无有效销量的花费保留为独立空烧行。货品销量及店铺利润使用剔除后结果。用户确认两个商品ID仅剔除商品、货品销售并调整店铺净利润，不扩展到店铺现金流。2026-11-09最新确认店铺支付/退款改按订单及售后计算，后台金额只做复核；访客、买家、评分等其他后台指标保持原口径。

用户要求更新已生成日报时，入口支持`--refresh-existing --database-root <本地库> --archive-dir <固定报表目录>`：备份后迁移已识别的历史表，保留原业务结果、date_notes、历史剔除及元数据，执行确认的新增剔除/暂计0、货品拆分与利润汇总，不重新导入原始订单。只把商品日报更新到原WPS在线表；货品和店铺日报本地交付。只修改货品金额/利润列时使用入口`--refresh-goods-only`，保留商品与店铺文件，不触发商品重复上传。

## 货品和店铺净利率（2026-11-09最新确认）

2026-11-10用户确认：货品名称中包含“临期”（兼容“【临期】”“临期-”等写法）的商品行及货品行，净利润固定计0，所有已生成日期与后续日期生效；按货品名称匹配，不凭SKU编码、标题或保质期推测。商品行含多个货品名称时，任一名称命中则该商品行净利润计0。销售额、销量、成本、实际毛利和推广费用保留，净利率仍按净利润÷净销售额计算，正分母为0%，非正/未知分母留空。备注记录此人工利润口径。店铺净利润按调整后的商品行汇总，净利率同步更新，其他店铺后台与现金流指标保留。更新历史通过`--refresh-existing`读取已确认日报，不重导原始订单。

货品月表末列增加净利率（第14列），保留原前13列顺序与手工透视表来源，净利率=货品净利润÷货品净销售额；现有“销售额”取去退款去补单后的实际成交金额，本身为净销售额，用该列作分母，不改金额或恢复旧补贴口径。店铺日报在店铺净利润后增加净利率（第28列），净利率=店铺净利润÷净成交额，分母沿用最新订单实收减当日成功退款的现金流口径。亏损保留负百分比；经确认暂计0的利润可计算0%；分母不大于0、利润或分母缺失时留空，不能伪造为0。按汇总后净利润除汇总后净销售额，不平均SKU净利率。百分比显示两位小数，Excel保留公式。

只为已生成报表补净利率时使用入口`--refresh-net-margins --archive-dir <固定报表目录>`，从现有货品/店铺最终数据计算，保留原金额、数量、排除结果、所有月份和店铺现金流/客户元数据，不重导原始订单，不更新商品XLSX或触发商品上传。货品14列、店铺28列，商品22列不变；店铺已识别26/27列档案可迁移。

## 店铺现金流（2026-11-09最新确认）

店铺支付金额及支付订单数以原始订单支付日统计，商家实收不因订单后来退款而先删除；退款按退款事件日另行扣减，包含更早支付订单的退款。退款只取“退款成功”的售后编号去重记录，当前按“同意退款时间”估算归日，不能用申请时间。后台金额仅作复核，差异写独立日志；表头说明估算，不写用户要求留空的异常说明。订单、售后快照按各自业务键保留最新记录，不能叠加多个导出。沿用用户明确的指定订单剔除，本地私有名单保存在数据库`_系统/日报配置/订单剔除.json`，不得发布到公开插件。两个商品ID销售排除不扩展到店铺现金流。已有月报可用入口`--refresh-store-cashflow`只更新支付、退款、净成交及相关客单价/ROI/费比；保存历史原值与原说明，替换旧净成交的人工保留值，商品/货品报表及店铺净利润不重算。细节及缺失处理见[现金流规则](references/store_cashflow.md)。

## Update Gate

This package includes the store-day extension. Read this installed SKILL.md before each request. Do not run the legacy two-skill updater automatically from a plugin cache: it does not update plugin installation metadata. Update a plugin through its marketplace, or a standalone skill through its installer. Keep local changes until an explicit update is requested.

## 店铺维度日报（本地扩展，2026-10-24）

默认同时输出原商品维度日报及独立的店铺维度日报。先读 [店铺规则](references/store_report_rules.md)。店铺使用 `assets/store_daily_template.xlsx` 的“①平台日报”，保留模板原26列并在末尾追加“店铺净利润、净利率”，按新版分区美化。店铺规则优先于模板内旧公式与文字；它不改变商品日报的订单/成本口径。店铺无需营销/成本表；商品日报受阻时仍可交付店铺已知字段并明确报告商品未生成。

入口仍为 `scripts/build_report.py`；正常默认更新两种固定报表。用户明确只要店铺时使用 `--store-only`。`--archive-dir` 统一指定固定报表目录；`--store-input <店铺日指标.csv或xlsx>` 可重复。`--output` 和 `--store-output` 仅在用户明确指定路径时使用，不能自动带上本次日期另建交付文件。

店铺金额按[订单与退款现金流](references/store_cashflow.md)计算：支付日汇总订单“商家实收金额(元)”，退款只计“退款成功”，按“同意退款时间”归日并明确标注估算；净成交额=当日订单实收−当日成功退款。后台金额仅作差异复核，不再作为支付、退款的填值来源。访客、买家、评分等店铺指标继续从本地后台来源或`--store-input`读取；金额缺订单/售后来源仍留空，不用后台兜底。标准输入映射和来源限制见店铺规则。自动读取有来源店铺元数据的“账户每日汇总”文件；不能叠加商品明细与账户汇总。

插件修改同步本地源与当前加载缓存，发布遵循文末作者更新流程。

## Mandatory Execution Order

### 商品日报生成后自动上传 WPS（2026-10-26）

每次商品日报成功生成并完成数据与版式校验后，默认将本次商品 XLSX 上传到当前用户 WPS 在线文档的“我的云文档”，返回可打开的在线文档链接。执行前读取 [WPS 上传与登录续传](references/wps_upload.md)。这属于用户已要求的常规日报交付步骤，无需每次再次确认上传。仅店铺日报、生成失败、待确认异常订单及内部 JSON/日志/预览不触发上传；店铺日报不在此自动上传范围内。用户当次明确要求仅本地时遵从用户。

登录、扫码、验证码交由用户操作；保留待上传文件与页面，用户完成后继续同一任务，不重算已完成的报表。上传异常与本地生成结果分别报告；不能把创建待上传记录或看到文件选择框当作已上传。入口会生成同名 `.wps-upload.json` 交接记录；代理必须执行浏览器上传及在线验证，Python 脚本自身不登录或上传。

Follow this order exactly. Do not skip a step, run a platform builder directly, or report completion before the final check.

1. Run the **Update Gate** and reread this file when the installed release is current or updated.
2. Start only through `scripts/build_report.py`; inspect order headers and select the platform automatically, or pass its `--platform` option only when the user explicitly requests one platform.
3. Read and validate the marketing-activity and cost workbooks before calculation. Stop and report the specific issue if a required marketing workbook cannot be read or lacks valid matching records.
4. Let the dispatcher run the internal platform builder. Do not invoke `build_tmall_report.py` or `build_pdd_report.py` as a separate user workflow.
5. Verify that every requested platform report was created, required output fields have been checked, and all warnings have been written to the generation log and sent in the current Codex or Work Buddy conversation.
6. 完成每个成功商品日报的 WPS 上传与在线验证，按 [WPS 上传与登录续传](references/wps_upload.md) 更新交接记录并交付链接；若登录或上传工具受阻，交付本地文件并准确报告待续传状态。

## Platform Routing

- **天猫**: Run `scripts/build_report.py --platform tmall`. Output `天猫店铺数据报表.xlsx`. When the order export has no `商家实收` field, use `买家应付` as `商家实收金额(元)` for receipt matching, sales statistics, small-receipt checks, and promotion allocation. Keep the current 天猫 rules: common Tmall export headers, and combine 万相台 `花费`、新客加速 `预估新客加速费用`、老客加速 `预估老客加速费用` (plus any legacy 品牌新享费用, which is included under 新客加速) into `推广费用`. In `备注`, show exactly three component amounts: `万相台费用`、`新客加速费用`、`老客加速费用`; 新客加速与老客加速均属于原品牌新享口径. Retain the zero-receipt registration-price adjustment for subsidy activities. Keep every promotion amount keyed by `日期 + 商品ID`, then preserve the existing rule of splitting that product’s fee across its matched 样式ID order rows by actual receipt amount; never allocate a grand total directly or evenly. Match costs using the order's specification-level merchant code and carry the matched cost table's `产线`、`项目组`、`管理类型`、`品种` into the output.
- **拼多多**: Run `scripts/build_report.py --platform pdd`. Output `数据报表拼多多.xlsx`. Keep the prior 拼多多 old-format rules, including style-ID aggregation, subsidy, cost, net-profit, and pricing reminders. Resolve every marketing field independently in this order: specification-safe `样式ID` → exact `商品ID + 商品SKU` → `店铺名称 + 规范化商品SKU` → cross-shop `商品SKU` candidates whose populated values are identical for that field. Normalize SKU full-width/half-width characters, all whitespace, and letter case. Never select the first cross-shop candidate when values conflict; a same-shop value wins before a cross-shop consistent value. A fallback record must not overwrite the order's original `商品ID`、`店铺名称` or final `商品SKU`. Keep safely matched promotion mechanisms as internal inputs; omit them from the public report. Activity and registration prices remain internal inputs for the existing order checks; do not output either price column. Populate `促销机制` from `活动名称`; only fall back to legacy promotion-mechanism fields when `活动名称` is blank. Normal-sales `到手价` is always actual receipt ÷ quantity. Compare it with `单件预估实收金额` only to generate warnings: even when actual is more than 5% below expected, write the actual value and warn. Never blank or replace the actual arrival price with the estimate.

For both platforms, never transfer a product's promotion spend to another selling product. For Pinduoduo, keep promotion spend keyed by `日期 + 店铺名称 + 商品ID` from import through allocation and reconciliation; allocate it only among effective sales rows for the same shop and product. When that key has no effective sales amount, add one independent `空烧推广费` row and preserve the promotion source shop. Keep every safely matched product, category, cost, and activity field; write both `实际成交数量（去退款去补单后）` and `实际成交金额（去退款去补单后）` as `0`, write gross profit as `0`, and write net profit as the negative promotion fee. If a product-level promotion row cannot resolve one specification, write `多规格未唯一匹配` in both `SKU ID` and `商品SKU`, never choose an arbitrary specification. Write `不适用（无有效销量）` in the empty-burn `到手价`. Mark `备注` as `空烧推广费：无有效销售`, append the multi-spec marker when applicable, retain the three Tmall promotion-component amounts in the Tmall remark, and record the generated empty-burn row in the generation log. Keep every normal sales row before every empty-burn row; place the complete empty-burn group at the absolute bottom of the output table, then sort that bottom group by `日期`、`店铺名称`、`商品ID`、`SKU ID`、`商品SKU`.

## 数据整理至日报交接保护

Apply these rules to files produced by the data-organizing Skill and to equivalent original exports:

1. Resolve each valid order row's business date independently in this priority: `订单支付时间` → `支付时间` → `订单付款时间` → `付款时间` → `订单成交时间`; for 天猫 only, finally fall back to `订单创建时间`. Use the first valid value on that row. If all available candidates are blank or invalid, stop and show the file, sheet, physical row, order number, and raw date values; never silently skip it or use the file name as the order date.
2. For a workbook produced by the data-organizing Skill, classify each promotion row first by `归档推广类型`, not by the combined workbook's header. Accept `拼多多-商品推广` only in the Pinduoduo builder; accept `万相台`、`新客加速` and `老客加速` only in the Tmall builder. Preserve each type's original fields when selecting its amount: 万相台 uses `花费`, 新客加速 uses `预估新客加速费用` or legacy 品牌新享 `预估抽佣金额`, and 老客加速 uses `预估老客加速费用`. For an original, unorganized export with no `归档推广类型`, detect platform and type by workbook headers as before. Skip a clearly recognized other-platform row; stop when a promotion-cost row lacks a valid type. Recognize legacy 品牌新享 from `业务日期` together with `预估抽佣金额`, and include it under the new-customer acceleration component.
3. Accept organizer keys `归档日期` and `归档商品ID` first; otherwise accept promotion date aliases `日期`、`业务日期`、`统计日期`、`统计时间`、`推广日期`、`报表日期` and product aliases `商品ID`、`商品id`、`商品Id`、`宝贝ID`、`主体ID`、`商品编号`、`商品编码`、`商品ID（必填）`.
4. Scan all workbook sheets, including hidden sheets, and locate eligible headers within the first 20 rows. Ignore instruction and auxiliary sheets that do not contain a valid required header. Exclude generated order summaries, organizing logs, missing-data reminders, interruption logs, backups, temporary files, and Excel lock files from order discovery.
5. Output exactly one text-formatted `SKU ID` column immediately after `商品ID`. Map 拼多多 `样式ID` and 天猫 `SKU ID`、`SKU_ID`、`SKUID` to it. Output the specification-level merchant code under `SKU编码`, followed by the cost-table `SKU名称`; never put a product name or title in the code. Keep different SKU IDs of the same product in separate rows.
6. When an order lacks a specification code, first use 商品数据's exact `商品ID + SKU ID → 规格编码` mapping, then the exact marketing-activity style record. Use a product-level mapping only when it resolves to exactly one specification. For an empty-burn row, fill `SKU ID` and `商品SKU` only when the product has exactly one resolvable specification; otherwise leave both blank. Match the cost table using the final output `商品SKU` code.
7. Before any sales, quantity, order-count, customer-price, product-ranking, promotion-allocation, or other operating calculation, classify every **order detail**. Exclude a detail when `订单性质=补差价` or `是否计销售=否`. For unmarked historical data, standardize `商品标题`/`商品名称` by trimming whitespace, normalizing full-width/half-width characters, and removing square-bracketed marketing prompts; exclude the detail when the standardized title contains `补差价` or `补收差价`. Never infer this from 商品ID、金额、数量 or SKU/specification. A mixed order must retain its ordinary product details. Do not write a separate 补差价 row in the daily workbook; record the excluded detail count and `已识别为补差价，不计销售` in the generation log and current Codex or Work Buddy conversation.

## Promotion Matching Protection

拼多多新版“商品营销”以报表列识别，账户每日汇总与商品分天明细均优先取“推广总花费(元)”，为空时取“推广成交花费(元)”。金额 0 是有效值。总营销花费、成交营销花费及结算券花费不计入日报推广费，优惠券已扣减商家实收。旧版商品推广沿用“总花费(元)”，为空时取“成交花费(元)”。全店托管仍按同日同店的账户推广费减商品推广费后分摊；商品、货品、店铺日报及相关利润统一使用此口径。

Apply all four safeguards before writing either platform workbook. A failure in any safeguard must stop report generation; never silently skip an invalid nonzero promotion row or continue with an unreconciled amount.

1. **Promotion backup gate**: Before reading any promotion amount, inspect every candidate file name and every parent directory under the promotion input roots. Treat names containing `备份`、`副本`、`backup`, or a standalone `bak`/`copy` marker, as backup candidates. If any candidate exists, stop before calculation, list the actual paths in the current conversation, and warn that multiple backups or snapshots can be interpreted as different shops and cause duplicate promotion spend. Ask the user to move them outside every promotion input root before rerunning. Never silently ignore them, reinterpret a backup directory as a shop, or continue by relying on file hashes, shop dimensions, snapshot deduplication, or input-output reconciliation.
2. **Daily date integrity**: Use a row-level `日期`、`业务日期`、`统计日期`、`统计时间`、`推广日期` or `报表日期` value when the promotion export provides one. Every nonzero-spend row must resolve to exactly one calendar day. Do not require one row for every calendar day in the requested range: a day with no promotion spend may be completely absent. For example, a 1–3 day export containing dated spend rows only for day 1 and day 3 is valid when day 2 had no spend. Never assign a multi-day export total to the range's first date. If the filename or date cell covers multiple days and the file has no single-day detail for every nonzero-spend row, stop and ask for a daily-detail export. This rule applies especially to Pinduoduo recent-seven-day promotion exports.
3. **Promotion snapshot deduplication**: First remove byte-for-byte duplicate files by SHA-256 and keep the newest copy. Then aggregate rows within each source file by `平台 + 店铺 + 推广类型 + 日期 + 商品ID`. Do not require, read, match, infer, or output a plan identity or plan ID; year/month archive folders are never plan identities. If the same complete business key occurs in multiple source files, keep only the newest file by modification time and write the ignored files to the generation log. Pinduoduo nonzero spend without a resolvable shop must stop instead of falling back to a cross-shop total.
4. **Input-output reconciliation**: Immediately before workbook writing, reconcile every deduplicated imported promotion amount against the sum of `推广费用` in all normal sales rows plus all `空烧推广费` rows. Reconcile Pinduoduo first by `日期 + 店铺名称 + 商品ID`; keep the existing Tmall `日期 + 商品ID` key. Then reconcile the platform total. The maximum permitted absolute difference is `0.01` yuan. If any key or the total exceeds this tolerance, stop, show the imported amount, output amount, and difference, and do not produce the workbook.

## Cost Matching and Data Quality Protection

### 单件综合成本与整行利润（2026-11-06用户最新确认）

商品日报成本列顺序为“综合成本、产品成本、快递成本、其他均摊成本、平台费用”，删除旧“店铺费用”和“平摊管理费用”。快递成本按下述运费表规则匹配，不再固定2.1元；其他均摊成本=到手价×16.7%；平台费用=到手价×0.7%；综合成本及四个分项都为每件金额，综合成本=产品成本+快递成本+其他均摊成本+平台费用。毛利=实际成交金额（去退款去补单后）−综合成本×实际成交数量（去退款去补单后）；毛利率=毛利÷实际成交金额（去退款去补单后），成交金额为0则留空。详见 [商品成本与毛利规则](references/product_economics.md)。计算与写表均按字段名定位，不再固定到旧列字母。SKU ID紧随商品ID，备注为最后一列。

商品日报不输出活动价、报名价、每单补贴金额、总补贴金额、净销售额（实际成交+总补贴金额），与2026-11-07确认删除的其他列一并移除后输出22列，SKU ID紧随商品ID，备注为最后一列。即使缺营销资料，可计算的综合成本、整行毛利、毛利率和净利润仍填写；删除字段不能由部分资料模式重新加回。历史列迁移只使用既有最终确认的报表行，保留date_notes、销量、实收、推广和排除结果；不得重导原始订单恢复已排除数据。

### 快递成本按编码与地区匹配（2026-11-06用户确认）

读取 `<数据库根目录>/1.产品数据库/运费成本.xlsx` 的“圆通各地运费”和“邮政各地运费”，按最终SKU编码精确匹配已保存报价，并使用有效订单的收货省市取地区列。黑龙江、吉林、辽宁固定取圆通；其他地区取圆通、邮政有效报价中的较低者，仅一家有效则取该家。黑吉辽圆通缺价时不能改取邮政。深圳使用深圳列；不使用“编码匹配”页当前所选地区替代订单地区。

同一报表行有多个地区时按有效商品数量加权，沿用每件成本口径；已剔除、退款无效及去重订单不能参与运费。2026-11-07用户补充：缺体积按重量核算；重量体积皆缺时参考同净克重商品的地区运费；省市缺失时采用各地区按上述快递规则选价后出现次数最多的报价，并列取较高价，不能用页面当前所选地区代替。估算必须记录来源，不补造实测重量、体积或地址；已知地址冲突仍留空。

同克重指整份销售单位的总净克重，组合装按可核实子件及数量求和，不能将单袋克重直接用于多袋套装。参考商品须有实测重量，不串联估算。先匹配同克重；仅CB10、LMB01、SQM05、WH03在无同克重参考时可用最接近克重产品，这是2026-11-07用户确认的限定例外。先按整份净克重的绝对差取最近组，距离并列取较高克重，再优先重量体积齐全、再优先同包装。逐快递、逐地区取参考报价中位数，沿用报价不按克重比例缩放；其他编码继续只用同克重。源表公式支持缺体积按实重；插件读取已保存报价，重量体积皆缺且无有效报价时按上述参考补算。无可用参考、编码缺失、地址冲突或数量核对不符时，快递及依赖利润留空。各报价已含0.40元/件打包费，不重复计提。最近克重估算记录目标克重、参考克重、差值和参考编码，不补造实测资料，不用于产品进货成本。历史回补只使用最终日报已确认的销量及成交额；订单集合未唯一核实时整行采用多数地区估算并标明，不恢复已剔除订单。详见 [商品成本与毛利规则](references/product_economics.md)。

### 净利润、日期和SKU列（2026-11-05再次确认）

净利润=该行毛利−该行推广费用。综合成本保持单件金额，只在毛利处乘一次有效销量，毛利和净利润不再重复乘销量，也不加入补贴。毛利未知时净利润留空。输出列“商品SKU”改为“SKU编码”，其后紧接新增“SKU名称”，按成本表的编码精确匹配“产品名称”（兼容旧表“商品名称”）。名称独立于成本金额匹配：成本不可用但编码名称明确时仍可填名称，不猜标题或组合子件；冲突停止。订单与营销输入仍兼容旧商品SKU字段；输出与历史业务键使用SKU编码。商品日期保存为真实日期，显示格式统一`yyyy-mm-dd`。SKU ID紧随商品ID，备注为最后一列。详见 [商品成本与毛利规则](references/product_economics.md)。

### 产品数据库成本（2026-11-05）

默认优先读取 `<数据库根目录>/1.产品数据库/产品成本.xlsx`，`--cost-table` 显式指定优先级最高；产品库无成本表时兼容旧的“营销活动监控”目录。读取新成本汇总格式和部分参考资料缺失时，先读 [产品成本读取规则](references/product_costs.md)。此规则覆盖下文仅从营销活动监控自动发现成本表的旧说明。

“编码成本汇总”按用户确认接入正式匹配成本和参考测算成本。状态为“已匹配”时优先使用正数“匹配成本(元)”，没有可用值时使用正数“参考测算成本(元)”；状态为“参考测算”时使用正数“参考测算成本(元)”，并在成本来源中标注。只有未匹配或没有可用正数成本的项目留空。组合装按组合装编码、货品按货品编码匹配订单最终规格编码；组合子件列表不作为组合装的规格编码。缺营销活动表时保留可用成本与按新定义可算的毛利、净利润，不输出已删除的补贴字段。

Apply these rules to both platforms while leaving the existing refund and supplemental-order handling unchanged:

1. Treat `SKU ID` as the platform's specification identity and the final order `商品SKU` as the specification-level merchant code; never replace it with a marketing-table display name before cost matching. The shared cost reader must support these equivalent cost-code headers: `商家编码-规格维度`、`商品SKU`、`商品编码`、`商品SKU名称`、`SKU名称` and `SKU编码`. If a table has more than one candidate header, select the first nonblank value **for that row** in this order: `商家编码-规格维度` → `商品SKU` → the remaining aliases. Do not skip a row merely because an earlier candidate cell is blank. Normalize full-width characters, all whitespace, and letter case before comparing; retain the actual header and raw code in the source log.
2. Read every cost-table worksheet up to the first `10,000` rows and `100` columns. Warn when a sheet exceeds this limit. Accept a positive numeric cost from `6.11成本价`、`成本价`、`产品成本` or `成本`; never convert a nonblank invalid value such as `待定` into zero. Stop if the cost workbook cannot be read, contains no usable positive cost, or has the same normalized cost code with different costs. Same-cost duplicate codes may be deduplicated only with an explicit warning that includes both worksheet rows and raw codes; never silently keep the last record.
3. Tmall keeps the cost-table-first order. For Pinduoduo, a valid marketing `产品成本` resolved by same-shop normalized SKU or by a cross-shop field-consistent SKU match overrides an older cost-table amount for the same final order SKU. Other Pinduoduo matches keep the cost-table-first behavior. Keep `项目组`、`管理类型`、`品种`、`产线` in the internal matched cost record even when a marketing cost overrides its amount; do not output these columns. If `样式ID` points to a different normalized SKU, skip that unsafe style match and emit a visible warning; never let it replace the final order SKU.
4. For every normal and empty-burn output row, write one structured record to the standalone `<日报文件名>_成本获取日志.csv` with `平台`、`行类型`、`日期`、`店铺名称`、`商品ID`、`样式ID`、`订单规格编码`、`产品成本`、exact `成本来源`、`营销成本来源`、`营销成本`、`成本表成本` and `成本差异（营销-成本表）`. A cost-table source must be recorded as `成本表:<文件名>/<工作表>/<实际编码表头>=<原始规格编码>`. When a Pinduoduo same-shop or cross-shop-consistent marketing cost overrides the cost table, also write a visible generation warning with both amounts and the difference. Mark the row type as `正常销售` or `空烧推广费`.
5. Treat an activity as a subsidy activity only when both prices are valid numbers and `报名价 > 活动价`. Calculate the per-unit subsidy strictly as `报名价 - 活动价`; use this same value for report calculations and the order cost-anomaly review. If `报名价 < 活动价`, do not calculate a negative subsidy; warn that the marketing data is probably wrong and identify the affected style, product, SKU, and prices. Do not determine subsidy status from promotion-mechanism text.
6. Deduplicate byte-for-byte identical order files, and when the same identified order line occurs across overlapping exports, keep the newest file by modification time using `店铺 + 订单号 + SKU ID + 商品ID + 商品SKU`. Do not merge unrelated rows merely because an order-number field is absent. Stop on an unreadable relevant order or promotion file, a required-field failure, or an invalid nonblank order/promotion number. Keep 商品ID、SKU ID and 商品SKU as text when writing Excel so long identifiers never lose digits.
7. For 商品数据 exports above `100,000` rows, use the first `100,000` rows and warn; do not discard the entire file.

For both platform outputs, display `实际成交数量（去退款去补单后）` as an integer, retained money columns with two decimal places, and `毛利率` as a percentage. The five deleted price/subsidy/net-sales fields are not report columns.

For both platforms, keep cost metadata in internal matching records. The public output ends with `备注`; `SKU ID` is immediately after `商品ID`.

Both modes use `%USERPROFILE%\Desktop\运营数据库`. Marketing activity workbooks are auto-detected from `营销活动监控`, preferring the most recently modified workbook whose name contains `营销活动_更新`. Cost discovery follows the product-database rule above, with the legacy folder as fallback. Only calculate dates explicitly requested by the user.

Read the marketing activity workbook before calculating. If it cannot be opened, has no `商品ID`、`样式ID`、`商品SKU` headers, or has no valid product records, stop immediately and report the specific failure. Never generate a report with blank marketing fields after a marketing-workbook read failure.

For Pinduoduo, check the retained `店铺名称`、`商品ID`、`SKU ID`、`SKU编码`、`产品成本` and `到手价` fields before writing. Internal marketing fields still support matching and validation; do not add deleted columns to the output. Never leave one of these fields silently blank. When it cannot be filled safely, write `营销活动表未匹配/未维护` and add a generation warning; use the empty-burn and multi-spec markers defined above where applicable. Preserve invalid nonblank values such as `#N/A` for diagnosis or replace an unusable cost with the explicit missing marker, but never coerce a missing or invalid cost or price to numeric zero.

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

Before producing either platform report, scan every eligible order and require confirmation when either condition is true: the order's original total actual receipt is greater than `0` and less than `1` yuan; or its subsidy-adjusted unit receipt is lower than the matched unit product cost by more than `10%`. Calculate the adjusted unit receipt as `(订单实收 ÷ 数量) + (报名价 - 活动价)`, but add the subsidy only when the matched marketing row has two valid numeric prices and `报名价 > 活动价`. Match the subsidy row by `样式ID` first, then exact `商品ID + 商品SKU`, and use a same-product fallback only when that product has exactly one marketing row; never select one of several ambiguous SKU rows. If the subsidy raises the adjusted receipt so that it is no longer more than `10%` below cost, do not show a confirmation reminder and do not interrupt report generation. The small-receipt check must continue to use the original actual receipt and must not be hidden by a subsidy.

Only review prices below cost; do not stop for prices above cost. Prefer the cost-table match by specification-level merchant/SKU code, then fall back to the marketing-activity product cost. If no positive cost can be matched, do not run the deviation comparison for that order; write a missing-cost warning to the generation log and send it in the current conversation. For a Tmall zero-receipt subsidy order that qualifies for the registration-price adjustment, use the registration-price-adjusted receipt as the complete comparison receipt and do not add the subsidy a second time. When a subsidized order remains more than `10%` below cost, the confirmation detail must show original unit receipt, per-unit subsidy, adjusted unit receipt, unit cost, and the adjusted deviation.

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

- `scripts/product_schema.py`: Canonical 22-column schema and recognized legacy header projection.
- `scripts/build_report.py`: Platform dispatcher.
- `scripts/build_tmall_report.py`: 天猫 calculation rules.
- `scripts/build_pdd_report.py`: 拼多多旧格式 calculation rules.
- `scripts/data_quality_protection.py`: Shared strict-number, cost-code normalization, duplicate-warning, price-validation, and standalone cost-acquisition log rules.
- `scripts/handoff_protection.py`: Shared order-date priority, cross-platform header classification, multi-sheet header search, file exclusion, and product/SKU mapping rules.
- `scripts/promotion_protection.py`: Shared promotion-backup gate, daily-date integrity, promotion-snapshot deduplication, and input-output reconciliation safeguards.
- `scripts/update-skills.ps1`: Check the shared public GitHub repository and transactionally update both Skills before business work.
- `references/template_columns.csv`: Shared output-column order.

## 本周数据及V3读取（本地2026-10-24补充）

商品与店铺日报均使用V3数据库；店铺扩展以references/store_report_rules.md末尾的补齐数据接入规则为准。用户已授权成本/营销缺失时先输出已知字段，可在入口增加`--allow-missing-reference`。默认仍严格校验；不隐藏未知字段，不伪造利润。商品推广实际和托管分摊必须分开核算。本周默认周一至昨日完整日。

## 新老客列补齐（2026-10-24用户确认）

店铺日报现填写新客数、老客数及后台老客占比。新老人数按已确认的ROUND规则推算，原始占比不反算覆盖；具体规则见references/store_report_rules.md末尾。
