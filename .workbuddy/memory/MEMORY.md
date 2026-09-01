# 项目长期记忆

## 项目定位
本机项目 `/Users/xiang/Documents/novel` 是一个中文长篇小说创作工作区，由 Obsidian（`.obsidian`）+ Syncthing（`.stfolder`/`.stignore`）+ Git 三件套管理。

## 目录约定
- `novel-project/设定/` — 世界观、文风基准、角色卡、时间线、势力关系
- `novel-project/大纲/` — 卷纲、章纲
- `novel-project/状态/` — 角色状态追踪、伏笔追踪表、道具持有表
- 正文目录按技能规范为 `novel-project/正文/`（当前尚未创建）

## 技能位置（重要）
小说创作技能 `novel-writing` 的原始文件在 `/Users/xiang/Documents/novel/.catpaw/skills/novel-writing/`（Catpaw 工具的位置）。
WorkBuddy 只扫描 `.workbuddy/skills/`，因此已建立软链接：
`.workbuddy/skills/novel-writing -> ../../.catpaw/skills/novel-writing`
**单一数据源，修改技能请改 `.catpaw` 下的原始文件，不要复制副本。**

技能含 8 个 md 文件：`SKILL.md`（主流程八阶段）、`style-benchmark.md`、`de-ai-humanizer.md`、`scene-card-crafter.md`、`critic-qa.md`、`reference.md`、`memory-tracker.md`、`examples.md`。

## 写作流程要点
一律走「动作分镜流 → 感官扩写」两步解耦，禁止单次生成同时推演剧情与雕琢文风；正文完成后必须过 `de-ai-humanizer.md` 禁用清单，再做双轨审稿与状态归档。
