# 基金季报 / 中报 / 年报监控

这是 `fund-monitor` 仓库中的第二套独立监控，不会覆盖原来的限购/限额监控。

## 监控内容

- 第一季度报告
- 半年度报告 / 中期报告
- 第三季度报告
- 年度报告

数据来自 AKShare 的东方财富基金公告接口。

## 运行频率

GitHub Actions 每小时的 07、17、27、37、47、57 分运行，也就是每 10 分钟扫描一次。

工作流：`.github/workflows/report-monitor.yml`

## 飞书

直接复用原限额监控已经使用的 GitHub Secret：

`FEISHU_WEBHOOK_URL`

默认消息关键词为：

`基金公告`

## 防重复逻辑

`report_state.json` 会记录：

- 哪些基金已经完成首次历史基线
- 哪些定期报告已经通知过

第一次扫描某只基金时，只把当前历史报告记入基线，不发送旧报告。以后只要发现新的季报、中报或年报，才发飞书。

## 基金池

基金列表位于：

`report_funds.json`

目前是 29 只 QDII 基金，与仓库旧监控中的 QDII 主基金池保持一致。

## 手动测试飞书

GitHub → Actions → fund-report-monitor → Run workflow

把 `send_test` 设为 `true`，会发送一条测试消息，不执行公告扫描。
