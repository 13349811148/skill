# AI Skill 自动同步仓库

这个仓库集中管理以下两个 Skill，并支持 Codex、腾讯 WorkBuddy 和 CodeBuddy：

- `shop-report-organizer`
- `data-report-multi-platform`

安装后使用中文显示名，并自动带上当前 `release.json` 的版本号：

- 日报 `<版本号>`
- 数据整理 `<版本号>`

Codex 使用 `agents/openai.yaml` 显示名称，WorkBuddy 使用 `_skillhub_meta.json` 显示名称。英文目录名和内部 ID 保持不变，以保证触发、校验和自动更新兼容。

每台电脑首次安装一次同步器。安装器会自动检测本机已有的软件，并写入对应目录：

- Codex：`%USERPROFILE%\.codex\skills`
- WorkBuddy：`%USERPROFILE%\.workbuddy\skills`
- CodeBuddy：`%USERPROFILE%\.codebuddy\skills`

两套 Skill 都内置更新程序。以后用户每次触发任意一个业务 Skill 时，Skill 会先从本仓库检查更新，再重新读取最新规则并继续当前任务。安装器还会尽可能配置 `SessionStart` 启动检查，作为额外保障。只有两套 Skill 都通过结构校验后，才会一起替换本机版本；下载、校验或替换失败时，本机旧版本会被保留，软件继续使用旧版本。

## 一、维护电脑准备

本仓库地址：<https://github.com/13349811148/skill>

仓库已经建立。维护电脑需要安装 Git 和 Python 3，并通过 GitHub 完成身份认证。推荐安装 GitHub CLI 后运行：

```text
gh auth login
gh auth status
```

首次取得仓库：

```text
git clone https://github.com/13349811148/skill.git
cd skill
```

如果本地目录已经是这个仓库，只需确认远端地址并拉取最新版本：

```text
git remote -v
git pull --ff-only
```

自动更新程序使用公开 GitHub 仓库下载，不要求同事安装 Git 或登录 GitHub。

## 二、其他电脑首次安装

现有旧版 Skill 本身没有更新程序，因此每台旧电脑仍需完成最后一次初始化。让同事打开 Windows PowerShell，粘贴下面这一条命令并回车：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://raw.githubusercontent.com/13349811148/skill/a7f8b90/bootstrap.ps1' | iex"
```

命令会自动下载临时安装文件、检测已安装的 Codex、WorkBuddy 或 CodeBuddy、完成配置并清理临时文件，不要求同事预装 Git。安装完成后，今后的 Skill 更新不再需要同事打开 PowerShell 或输入命令。

也可以下载或克隆本仓库后，直接双击 `install-sync.cmd`。

也可以在仓库目录手工运行：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\install-skill-sync.ps1
```

安装完成后完全退出并重新打开 Codex、WorkBuddy 或 CodeBuddy。首次启动如果出现钩子审核或信任提示，请审核并信任 `update-skills.ps1`。

## 三、跨平台发布更新

发布器支持 Windows、macOS 和 Linux。它会自动查找以下位置中同时包含两套 Skill 的第一个目录：

1. `$CODEX_HOME/skills`
2. `$HOME/.agents/skills`
3. `$HOME/.codex/skills`
4. `$HOME/.workbuddy/skills`
5. `$HOME/.codebuddy/skills`

如果多处都安装了 Skill，请用 `--skills-dir` 或对应包装脚本的 `-SkillsDir` 明确指定来源，避免发布旧副本。

先更新并测试本机安装的两套 Skill，然后拉取仓库最新版本：

```text
git pull --ff-only
```

### macOS 或 Linux

推荐先创建本地提交但不推送，检查无误后再推送：

```sh
./publish-skills.sh --message "说明本次更新内容" --no-push
git show --stat --name-status HEAD
git push -u origin HEAD
```

也可以一步完成提交和推送：

```sh
./publish-skills.sh --message "说明本次更新内容"
```

如果要明确使用 Codex 兼容目录：

```sh
./publish-skills.sh \
  --skills-dir "$HOME/.codex/skills" \
  --message "说明本次更新内容"
```

### Windows PowerShell

推荐先本地提交并检查：

```powershell
.\publish-skills.ps1 `
  -SkillsDir "$env:USERPROFILE\.codex\skills" `
  -Message "说明本次更新内容" `
  -NoPush
git show --stat --name-status HEAD
git push -u origin HEAD
```

也可以一步完成：

```powershell
.\publish-skills.ps1 -Message "说明本次更新内容"
```

Windows 还可以直接双击 `publish-skills.cmd`。Python 核心发布器也可以在所有系统直接运行：

```text
python3 publish-skills.py --message "说明本次更新内容"
```

`--no-push`/`-NoPush` 仍会生成本地 Git 提交，但不会推送。发布器拒绝覆盖已暂存的 Git 变更；发布前请先提交或取消暂存。脚本会把本机已安装的两套 Skill 事务式同步进仓库、生成新版本号、更新显示元数据、提交并按需推送。推送失败时，本地提交会保留，之后可以直接运行 `git push -u origin HEAD`。

其他电脑下一次触发任意一个业务 Skill 时会自动更新；支持启动钩子的客户端也会在新建或恢复任务时提前检查。

## 更新顺序

1. 用户触发任意一个业务 Skill，内置更新程序先运行一次。
2. 更新程序从 GitHub 下载目标分支的最新发布包。
3. 在临时目录解包并校验 `release.json`、Skill 目录及 `SKILL.md` 名称。
4. 两套 Skill 作为一个事务替换；任何一步失败都会回退到旧版本。
5. Skill 重新读取已更新的 `SKILL.md`，然后执行当前业务规则。
6. 支持 `SessionStart` 钩子的客户端还会在任务开始时提前执行同样的检查。
