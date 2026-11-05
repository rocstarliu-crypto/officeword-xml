# Word 实际 ECS 功能配置验证 V1.9

V1.8 在当前 Word 实际缓存中发现完整功能名，本版测试 UCI.ShowFloatieChangeGate、UXPlatform 的四个 CopilotDAB 开关，以及实际存在的 Word/Shared Copilot 相关开关；这些与此前根据短标识推测命名空间的候选不同。仍要求两次无弹窗且复制正确、恢复后故障重现、只读与隐私禁用策略不变。

仅专用标签显式触发；启动前检查滚动时间窗口内流程和任务数均不超过10，单一工作流排队。
