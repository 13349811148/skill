#!/usr/bin/env python3
"""Publish the installed Skill pair from Windows, macOS, or Linux."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence


EXPECTED_GITHUB_OWNER = "13349811148"
EXPECTED_GITHUB_REPOSITORY = "skill"
EXPECTED_GITHUB_FULL_NAME = f"{EXPECTED_GITHUB_OWNER}/{EXPECTED_GITHUB_REPOSITORY}"
LOCAL_ONLY_EXIT_CODE = 3
SKILL_NAMES = (
    "shop-report-organizer",
    "data-report-multi-platform",
)
DISPLAY_LABELS = {
    "data-report-multi-platform": "日报",
    "shop-report-organizer": "数据整理",
}
PUBLISH_PATHS = (
    "release.json",
    "skills",
    "sync",
    "bootstrap.ps1",
    "install-skill-sync.ps1",
    "install-sync.cmd",
    "publish-skills.py",
    "publish-skills.sh",
    "publish-skills.ps1",
    "publish-skills.cmd",
    "README.md",
    ".gitignore",
)


class PublishError(RuntimeError):
    """A user-facing publishing failure."""


class LocalOnlyPublishError(PublishError):
    """Cloud publication is blocked while installed Skill edits stay local."""


def run_git(
    repository: Path,
    arguments: Sequence[str],
    *,
    capture: bool = False,
    allowed_codes: Iterable[int] = (0,),
) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        ["git", "-C", str(repository), *arguments],
        check=False,
        text=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
    )
    if completed.returncode not in set(allowed_codes):
        detail = (completed.stderr or completed.stdout or "").strip()
        command = "git " + " ".join(arguments)
        raise PublishError(f"{command} failed ({completed.returncode}): {detail}")
    return completed


def run_gh(
    arguments: Sequence[str],
    *,
    capture: bool = False,
    allowed_codes: Iterable[int] = (0,),
) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        ["gh", *arguments],
        check=False,
        text=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
    )
    if completed.returncode not in set(allowed_codes):
        detail = (completed.stderr or completed.stdout or "").strip()
        command = "gh " + " ".join(arguments)
        raise PublishError(f"{command} failed ({completed.returncode}): {detail}")
    return completed


def github_full_name_from_remote(remote_url: str) -> str | None:
    value = remote_url.strip()
    patterns = (
        r"^https?://github\.com/([^/]+)/([^/]+?)(?:\.git)?/?$",
        r"^git@github\.com:([^/]+)/([^/]+?)(?:\.git)?/?$",
        r"^ssh://git@github\.com/([^/]+)/([^/]+?)(?:\.git)?/?$",
    )
    for pattern in patterns:
        match = re.match(pattern, value, flags=re.IGNORECASE)
        if match:
            return f"{match.group(1)}/{match.group(2)}"
    return None


def local_only_publish_error(reason: str) -> LocalOnlyPublishError:
    return LocalOnlyPublishError(
        "PUBLISH_BLOCKED_LOCAL_ONLY\n"
        f"作者身份校验未通过：{reason}\n"
        "本地 Skill 修改可以保留，并且仅对当前电脑生效；云端仓库没有更新。\n"
        "如需更新云端版本，请联系仓库作者。请让当前 AI 根据本次本地修改，"
        "生成一段可复制给作者的中文更新提示词。"
    )


def read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise PublishError(f"Cannot read JSON file {path}: {exc}") from exc


def write_bytes_atomic(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    file_mode = stat.S_IMODE(path.stat().st_mode) if path.exists() else 0o644
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.chmod(temporary_name, file_mode)
        os.replace(temporary_name, path)
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise


def write_text_atomic(path: Path, text: str) -> None:
    write_bytes_atomic(path, text.encode("utf-8"))


def write_json_atomic(path: Path, value: dict) -> None:
    write_text_atomic(path, json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def validate_skill(skill_directory: Path, expected_name: str) -> None:
    manifest_path = skill_directory / "SKILL.md"
    if not manifest_path.is_file():
        raise PublishError(f"Installed Skill is missing: {manifest_path}")

    try:
        skill_text = manifest_path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        raise PublishError(f"Cannot read {manifest_path}: {exc}") from exc

    match = re.search(
        r"(?ms)^\s*---\s*$.*?^\s*name:\s*[\"']?([a-z0-9-]+)", skill_text
    )
    if match is None or match.group(1) != expected_name:
        raise PublishError(f"SKILL.md validation failed for {expected_name}")

    for path in skill_directory.rglob("*"):
        if path.is_symlink():
            raise PublishError(f"Symlinks are not published automatically: {path}")


def contains_all_skills(skills_directory: Path) -> bool:
    return all((skills_directory / name / "SKILL.md").is_file() for name in SKILL_NAMES)


def source_candidates() -> list[Path]:
    candidates: list[Path] = []
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        candidates.append(Path(codex_home).expanduser() / "skills")

    user_home = Path.home()
    candidates.extend(
        (
            user_home / ".agents" / "skills",
            user_home / ".codex" / "skills",
            user_home / ".workbuddy" / "skills",
            user_home / ".codebuddy" / "skills",
        )
    )

    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = os.path.normcase(str(candidate.resolve(strict=False)))
        if key not in seen:
            unique.append(candidate)
            seen.add(key)
    return unique


def resolve_skills_directory(
    explicit_skills_directory: str | None, explicit_client_home: str | None
) -> Path:
    if explicit_skills_directory:
        skills_directory = Path(explicit_skills_directory).expanduser().resolve()
        if not contains_all_skills(skills_directory):
            raise PublishError(
                f"The selected skills directory does not contain both Skills: {skills_directory}"
            )
        return skills_directory

    if explicit_client_home:
        skills_directory = Path(explicit_client_home).expanduser().resolve() / "skills"
        if not contains_all_skills(skills_directory):
            raise PublishError(
                f"The selected client home does not contain both Skills: {skills_directory}"
            )
        return skills_directory

    for candidate in source_candidates():
        if contains_all_skills(candidate):
            return candidate.resolve()

    checked = "\n  - ".join(str(path) for path in source_candidates())
    raise PublishError(
        "No installed source copy of both Skills was found. Checked:\n  - " + checked
    )


def copy_skill(source: Path, destination: Path) -> None:
    def ignore(_directory: str, names: list[str]) -> set[str]:
        return {
            name
            for name in names
            if name == "__pycache__" or name.lower().endswith(".pyc")
        }

    shutil.copytree(source, destination, ignore=ignore)


class SkillSwap:
    """Atomically replace repository Skill folders with rollback support."""

    def __init__(self, repository: Path, transaction_root: Path) -> None:
        self.repository = repository
        self.transaction_root = transaction_root
        self.backup_root = transaction_root / "old"
        self.staged_root = transaction_root / "new"
        self.applied = False
        self.accepted = False

    def prepare(self, source_root: Path) -> None:
        self.staged_root.mkdir(parents=True)
        self.backup_root.mkdir(parents=True)
        for name in SKILL_NAMES:
            source = source_root / name
            validate_skill(source, name)
            staged = self.staged_root / name
            copy_skill(source, staged)
            validate_skill(staged, name)

    def apply(self) -> None:
        destination_root = self.repository / "skills"
        destination_root.mkdir(parents=True, exist_ok=True)
        moved_names: list[str] = []
        try:
            for name in SKILL_NAMES:
                destination = destination_root / name
                backup = self.backup_root / name
                had_destination = destination.exists()
                if destination.exists():
                    destination.rename(backup)
                try:
                    (self.staged_root / name).rename(destination)
                except Exception:
                    if had_destination and backup.exists():
                        backup.rename(destination)
                    raise
                moved_names.append(name)
        except Exception:
            for name in reversed(moved_names):
                destination = destination_root / name
                backup = self.backup_root / name
                if destination.exists():
                    shutil.rmtree(destination)
                if backup.exists():
                    backup.rename(destination)
            raise
        self.applied = True

    def accept(self) -> None:
        self.accepted = True

    def rollback(self) -> None:
        if not self.applied or self.accepted:
            return
        destination_root = self.repository / "skills"
        for name in reversed(SKILL_NAMES):
            destination = destination_root / name
            backup = self.backup_root / name
            if destination.exists():
                shutil.rmtree(destination)
            if backup.exists():
                backup.rename(destination)
        self.applied = False


def set_skill_display_metadata(
    skill_directory: Path, internal_name: str, display_name: str, version: str
) -> None:
    metadata_path = skill_directory / "agents" / "openai.yaml"
    if not metadata_path.is_file():
        raise PublishError(f"Skill UI metadata is missing: {metadata_path}")

    lines = metadata_path.read_text(encoding="utf-8-sig").splitlines()
    indexes = [
        index
        for index, line in enumerate(lines)
        if re.match(r"^\s*display_name:\s*", line)
    ]
    if len(indexes) != 1:
        raise PublishError(f"Expected exactly one display_name in {metadata_path}")

    index = indexes[0]
    indentation = re.match(r"^\s*", lines[index]).group(0)  # type: ignore[union-attr]
    escaped_name = display_name.replace("\\", "\\\\").replace('"', '\\"')
    lines[index] = f'{indentation}display_name: "{escaped_name}"'
    write_text_atomic(metadata_path, "\n".join(lines) + "\n")

    skillhub_metadata = {
        "name": display_name,
        "slug": internal_name,
        "version": version,
    }
    write_json_atomic(skill_directory / "_skillhub_meta.json", skillhub_metadata)


def next_release_timestamp(previous_version: str | None) -> tuple[str, str]:
    current = datetime.now().astimezone()
    release_version = current.strftime("%Y.%m.%d.%H%M%S")
    if previous_version == release_version:
        time.sleep(1.05)
        current = datetime.now().astimezone()
        release_version = current.strftime("%Y.%m.%d.%H%M%S")
    return release_version, current.isoformat(timespec="microseconds")


def ensure_git_context(repository: Path) -> None:
    if shutil.which("git") is None:
        raise PublishError("Git is required and was not found in PATH.")

    top_level = run_git(
        repository, ("rev-parse", "--show-toplevel"), capture=True
    ).stdout.strip()
    if Path(top_level).resolve() != repository.resolve():
        raise PublishError(f"The publisher must run from the repository root: {repository}")

    unresolved = run_git(
        repository,
        ("diff", "--name-only", "--diff-filter=U"),
        capture=True,
    ).stdout.strip()
    if unresolved:
        raise PublishError("Resolve merge conflicts before publishing:\n" + unresolved)

    staged = run_git(
        repository,
        ("diff", "--cached", "--quiet"),
        allowed_codes=(0, 1),
    )
    if staged.returncode == 1:
        raise PublishError(
            "The Git index already contains staged changes. "
            "Commit or unstage them before publishing."
        )


def ensure_author_publisher(repository: Path) -> None:
    """Require the repository owner account before any release mutation."""
    remote_url = run_git(
        repository, ("remote", "get-url", "origin"), capture=True
    ).stdout.strip()
    remote_full_name = github_full_name_from_remote(remote_url)
    if remote_full_name is None:
        raise local_only_publish_error(f"无法识别 origin：{remote_url or '未配置'}")
    if remote_full_name.casefold() != EXPECTED_GITHUB_FULL_NAME.casefold():
        raise local_only_publish_error(
            f"origin 指向 {remote_full_name}，要求为 {EXPECTED_GITHUB_FULL_NAME}"
        )

    if shutil.which("gh") is None:
        raise local_only_publish_error("未安装 GitHub CLI，无法验证作者账号")
    try:
        run_gh(("auth", "status", "--hostname", "github.com"), capture=True)
        login = run_gh(("api", "user", "--jq", ".login"), capture=True).stdout.strip()
    except PublishError as exc:
        raise local_only_publish_error(f"GitHub CLI 未登录或认证失败：{exc}") from exc
    if login != EXPECTED_GITHUB_OWNER:
        raise local_only_publish_error(
            f"当前 GitHub 账号为 {login or '未识别'}，要求为 {EXPECTED_GITHUB_OWNER}"
        )

    try:
        permission = run_gh(
            (
                "repo",
                "view",
                EXPECTED_GITHUB_FULL_NAME,
                "--json",
                "viewerPermission",
                "--jq",
                ".viewerPermission",
            ),
            capture=True,
        ).stdout.strip()
    except PublishError as exc:
        raise local_only_publish_error(f"无法读取目标仓库权限：{exc}") from exc
    if permission != "ADMIN":
        raise local_only_publish_error(
            f"账号 {login} 对 {EXPECTED_GITHUB_FULL_NAME} 的权限为 {permission or '未知'}，要求为 ADMIN"
        )

    print(
        f"作者身份校验通过：{login}，仓库 {EXPECTED_GITHUB_FULL_NAME}，权限 {permission}。",
        flush=True,
    )


def ensure_git_identity(repository: Path) -> None:
    missing: list[str] = []
    for key in ("user.name", "user.email"):
        value = run_git(
            repository,
            ("config", "--get", key),
            capture=True,
            allowed_codes=(0, 1),
        ).stdout.strip()
        if not value:
            missing.append(key)
    if missing:
        raise PublishError(
            "Git author identity is not configured. Run:\n"
            '  git config --global user.name "Your Name"\n'
            '  git config --global user.email "you@example.com"'
        )


def git_status(repository: Path, paths: Sequence[str]) -> str:
    return run_git(
        repository,
        ("status", "--porcelain", "--", *paths),
        capture=True,
    ).stdout.strip()


def publish(args: argparse.Namespace) -> int:
    repository = Path(__file__).resolve().parent
    ensure_git_context(repository)
    ensure_author_publisher(repository)
    source_root = resolve_skills_directory(args.skills_dir, args.client_home)
    print(f"Skill source: {source_root}", flush=True)

    manifest_path = repository / "release.json"
    manifest_original = manifest_path.read_bytes()

    with tempfile.TemporaryDirectory(
        prefix=".publish-staging-", dir=repository
    ) as temporary_directory:
        swap = SkillSwap(repository, Path(temporary_directory))
        swap.prepare(source_root)
        swap.apply()

        try:
            skill_changes = git_status(repository, ("skills",))
            if not skill_changes:
                swap.accept()
                print("两套 Skill 与仓库版本一致，没有需要发布的改动。")
                return 0

            manifest = read_json(manifest_path)
            if int(manifest.get("schema_version", 0)) != 1:
                raise PublishError(
                    f"Unsupported release.json schema version: {manifest.get('schema_version')}"
                )
            listed_skills = set((manifest.get("skills") or {}).keys())
            if listed_skills != set(SKILL_NAMES):
                raise PublishError("release.json does not list the expected Skill pair.")

            release_version, published_at = next_release_timestamp(
                str(manifest.get("release_version") or "")
            )
            manifest["release_version"] = release_version
            manifest["published_at"] = published_at
            write_json_atomic(manifest_path, manifest)

            destination_root = repository / "skills"
            for name, label in DISPLAY_LABELS.items():
                set_skill_display_metadata(
                    destination_root / name,
                    name,
                    f"{label} {release_version}",
                    release_version,
                )

            ensure_git_identity(repository)
            run_git(repository, ("add", "--", *PUBLISH_PATHS))
            commit_message = args.message or f"Update Skills {release_version}"
            run_git(repository, ("commit", "-m", commit_message))
            swap.accept()

            if not args.no_push:
                try:
                    run_git(repository, ("push", "-u", "origin", "HEAD"))
                except PublishError as exc:
                    raise PublishError(
                        f"git push failed. The commit remains available locally. {exc}"
                    ) from exc

            action = "已生成本地发布提交" if args.no_push else "已发布"
            print(f"{action}版本 {release_version}。")
            return 0
        except Exception:
            if not swap.accepted:
                run_git(
                    repository,
                    ("reset", "--", *PUBLISH_PATHS),
                    capture=True,
                    allowed_codes=(0, 1),
                )
                write_bytes_atomic(manifest_path, manifest_original)
                swap.rollback()
            raise


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync, version, commit, and publish the two installed Skills."
    )
    parser.add_argument("-m", "--message", help="Git commit message")
    parser.add_argument(
        "--client-home",
        "--codex-home",
        dest="client_home",
        help="Client home containing a skills directory (for example ~/.codex)",
    )
    parser.add_argument(
        "--skills-dir",
        help="Exact directory containing both Skill folders; overrides auto-detection",
    )
    parser.add_argument(
        "--no-push",
        action="store_true",
        help="Create and verify the local commit without pushing it",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    try:
        return publish(parse_args(argv if argv is not None else sys.argv[1:]))
    except LocalOnlyPublishError as exc:
        print(str(exc), file=sys.stderr)
        return LOCAL_ONLY_EXIT_CODE
    except PublishError as exc:
        print(f"发布失败：{exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("发布已取消。", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"发布失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
