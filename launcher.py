# -*- coding: utf-8 -*-
"""
exe 런처 — PyInstaller로 빌드된 견적 표준 양식 변환기 실행 파일의 진입점.

동작:
  1. 번들된 소스(app.py, config.py, parsers/, converters/, classification/,
     templates/, data/)를 exe 옆 디렉터리로 복사
     → config.py의 BASE_DIR = Path(__file__).parent 가 exe 옆이 되어
       templates/ data/ 가 exe 옆에서 찾히도록
  2. Streamlit 서버를 현재 프로세스에서 실행 (LAN 접속 가능)
  3. 콘솔에 접속 주소 표시 (Streamlit이 Network URL 출력)
"""
import os
import shutil
import socket
import subprocess
import sys
import time
from pathlib import Path

PORT = 8501


def _kill_port(port: int) -> None:
    """포트에 리스닝 중인 프로세스 종료 (자기 자신 제외). Windows 전용."""
    if os.name != "nt":
        return
    try:
        out = subprocess.check_output(
            ["netstat", "-ano"], text=True, errors="ignore"
        )
    except Exception:
        return
    pids = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) < 5:
            continue
        if parts[0].startswith("TCP") and parts[3] == "LISTENING":
            if parts[1].endswith(f":{port}"):
                pid = parts[4]
                if pid.isdigit() and int(pid) != os.getpid():
                    pids.add(int(pid))
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
    if pids:
        time.sleep(2)


def _port_in_use(port: int) -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("0.0.0.0", port))
        return False
    except OSError:
        return True
    finally:
        s.close()


def _app_dir() -> Path:
    """exe(또는 개발 시 launcher.py)가 있는 디렉터리"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _bundle_dir() -> Path:
    """PyInstaller --onefile 번들 리소스 디렉터리 (개발 시 = 앱 디렉터리)"""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _safe_copytree(src: Path, dst: Path) -> None:
    """잠긴 파일을 건너뛰며 디렉터리 복사"""
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        target = dst / item.name
        try:
            if item.is_dir():
                _safe_copytree(item, target)
            else:
                shutil.copy2(item, target)
        except (PermissionError, OSError):
            pass  # 잠긴 파일은 기존 파일 유지


def _copy_bundle() -> None:
    app_dir = _app_dir()
    bundle = _bundle_dir()
    for name in ("app.py", "config.py", "run_conversion.py",
                 "parsers", "converters", "classification",
                 "templates", "data"):
        src = bundle / name
        dst = app_dir / name
        if not src.exists():
            continue
        if src.is_dir():
            # 기존 디렉터리 삭제 (잠긴 파일은 무시)
            if dst.exists():
                shutil.rmtree(dst, ignore_errors=True)
            _safe_copytree(src, dst)
        else:
            try:
                shutil.copy2(src, dst)
            except (PermissionError, OSError):
                # 잠긴 파일은 기존 파일 유지
                pass


def _ensure_config_toml(app_dir: Path) -> None:
    """frozen exe에서 developmentMode=false를 강제하는 config.toml 생성.

    PyInstaller --onefile 환경에서는 streamlit의 __file__이 _MEIPASS에 있어
    "site-packages"가 포함되지 않아 developmentMode가 True로 판정된다.
    이 모드에서는:
      - server.port 플래그 사용 불가 (RuntimeError)
      - 정적 에셋(프론트엔드) 라우트 미등록 → 루트 404
    config.toml로 developmentMode=false를 설정하면 정상 프로덕션 모드로 동작.
    """
    streamlit_dir = app_dir / ".streamlit"
    streamlit_dir.mkdir(parents=True, exist_ok=True)
    toml_path = streamlit_dir / "config.toml"
    content = "[global]\ndevelopmentMode = false\n"
    # BOM 없는 UTF-8로 작성 (PowerShell Set-Content -Encoding UTF8은 BOM 포함)
    toml_path.write_text(content, encoding="utf-8")


def main() -> None:
    # 1) 기존 서버 종료 (재실행 시 파일 잠금 방지)
    if _port_in_use(PORT):
        _kill_port(PORT)

    # 2) 번들 리소스 복사
    _copy_bundle()
    app_dir = _app_dir()
    os.chdir(app_dir)

    # 3) developmentMode=false 강제 (frozen exe용)
    _ensure_config_toml(app_dir)

    # 4) Streamlit 서버를 현재 프로세스에서 실행 (콘솔에 Network URL 출력)
    #
    # 주의: frozen(exe) 환경에서는 __file__에 "site-packages"가 없어
    # global.developmentMode가 True로 판정된다. 이 모드에서는 server.port를
    # 명시적으로 지정하면 "server.port does not work when global.developmentMode
    # is true" 오류가 난다. Streamlit 기본 포트가 이미 8501이므로 --server.port
    # 플래그를 주지 않고 기본값(8501)을 사용한다.
    sys.argv = [
        "streamlit", "run", str(app_dir / "app.py"),
        "--server.address", "0.0.0.0",
        "--server.headless", "true",
    ]
    from streamlit.web import cli
    cli.main()


if __name__ == "__main__":
    main()
