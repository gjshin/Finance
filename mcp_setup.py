"""Claude 데스크톱 앱에 DART 서버를 등록한다.

run_mcp_setup.bat 이 호출한다. 직접 실행할 수도 있다:
    python mcp_setup.py <API키> <파이썬경로> <dart_mcp.py경로>

기존 설정 파일은 손대기 전에 백업한다.
"""

import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

SERVER_KEY = "opendart"


def config_path():
    """OS별 Claude 데스크톱 설정 파일 위치"""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA")
        if not base:
            return None
        return Path(base) / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def build_entry(python_exe, server_py, api_key):
    return {
        "command": str(python_exe),
        "args": [str(server_py)],
        "env": {"OPENDART_API_KEY": api_key},
    }


def install(api_key, python_exe, server_py, path=None):
    """설정 파일에 서버를 등록하고 (성공여부, 메시지, 백업경로)를 돌려준다."""
    path = Path(path) if path else config_path()
    if path is None:
        return False, "설정 파일 위치를 찾지 못했습니다.", None

    existing = {}
    backup = None
    if path.exists():
        try:
            existing = json.loads(path.read_text(encoding="utf-8") or "{}")
        except json.JSONDecodeError:
            return False, f"기존 설정 파일이 손상되어 있습니다: {path}", None
        # 남의 설정을 날리지 않도록 먼저 복사해 둔다
        backup = path.with_suffix(f".json.backup-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(path, backup)

    servers = existing.setdefault("mcpServers", {})
    replaced = SERVER_KEY in servers
    servers[SERVER_KEY] = build_entry(python_exe, server_py, api_key)

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8")

    others = [k for k in servers if k != SERVER_KEY]
    msg = ("기존 opendart 설정을 새로 덮어썼습니다." if replaced
           else "opendart 서버를 새로 등록했습니다.")
    if others:
        msg += f" (다른 서버 {len(others)}개는 그대로 두었습니다: {', '.join(others)})"
    return True, msg, backup


def main():
    if len(sys.argv) < 4:
        print("사용법: python mcp_setup.py <API키> <파이썬경로> <dart_mcp.py경로>")
        return 1

    api_key, python_exe, server_py = sys.argv[1], sys.argv[2], sys.argv[3]
    ok, msg, backup = install(api_key, python_exe, server_py)

    print()
    if not ok:
        print(f"  [실패] {msg}")
        print()
        print("  아래 내용을 설정 파일에 직접 넣으셔야 합니다:")
        print(json.dumps({"mcpServers": {SERVER_KEY: build_entry(python_exe, server_py, api_key)}},
                         indent=2, ensure_ascii=False))
        return 1

    print(f"  [완료] {msg}")
    print(f"         설정 파일: {config_path()}")
    if backup:
        print(f"         이전 설정 백업: {backup.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
