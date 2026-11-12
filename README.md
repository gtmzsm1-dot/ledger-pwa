# ledger-pwa

离线记账网页应用（PWA）。纯前端、无后端、无统计脚本；数据只保存在本机浏览器存储（IndexedDB），页面通过 CSP 禁止任何外部连接。

- 金额一律以整数“分”计算
- 加到 iPhone 主屏幕后离线可用
- 第三方库：SheetJS 社区版（vendor/xlsx.mjs，Apache-2.0，见 vendor/xlsx.LICENSE）
