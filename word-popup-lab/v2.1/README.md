# 连接体验提示调用时机验证 V2.1

V2.0 在真实ECS配置中发现完整标识 Microsoft.Office.UXPlatform.RestrictSUAToLeftClick，并在 Word 二进制发现查看模式与 SDX 启动、回调、只读选择提示有关的标识。本版对这些候选做正反及组合测试。必须同时通过普通/结构副本、键盘/鼠标、重启及恢复对照，并保持页面视图、同一路径、文件只读和 Document.ReadOnly=True、连接体验全部关闭。
