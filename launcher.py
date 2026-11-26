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


def _kill_other_instances() -> None:
    """자기 자신 이름의 다른 exe 인스턴스 종료.

    이전 인스턴스가 templates/*.xlsx 등 파일을 잠그고 있으면 새 인스턴스가
    덮어쓸 수 없다(PermissionError). 포트 8501에 바인딩되기 전이거나
    파일 핸들 해제에 시간이 걸리는 경우를 대비해, 같은 이름의 다른
    프로세스를 모두 종료한다.
    """
    if os.name != "nt":
        return
    # frozen exe의 프로세스 이름 = exe 파일명(확장자 제외)
    exe_name = Path(sys.executable).stem if getattr(sys, "frozen", False) else None
    if not exe_name:
        return
    try:
        out = subprocess.check_output(
            ["tasklist", "/FO", "CSV", "/NH"], text=True, errors="ignore"
        )
    except Exception:
        return
    my_pid = os.getpid()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split(",")]
        if len(parts) < 2:
            continue
        if parts[0] == exe_name and parts[1].isdigit() and int(parts[1]) != my_pid:
            try:
                subprocess.run(["taskkill", "/F", "/PID", parts[1]],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
    time.sleep(2)


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


def _dbg(msg: str) -> None:
    try:
        with open(Path(os.environ.get("TEMP", ".")) / "launcher_debug.log", "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def _copy_with_retry(src: Path, dst: Path, retries: int = 8, delay: float = 1.0) -> bool:
    """잠긴 파일을 재시도하며 복사. 성공 시 True.

    Windows Defender 등 AV가 Office 문서(xlsx) 최종 파일명 쓰기를 실시간
    스캔하면서 잠그는 경우를 우회하기 위해, 같은 디렉터리에 임시 파일명으로
    쓴 뒤 os.replace로 최종 이름으로 바꾼다.
    """
    for attempt in range(retries):
        tmp = dst.with_name(dst.name + f".tmp{attempt}")
        try:
            # 바이트만 복사 (메타데이터/읽기전용 속성 미전파)
            shutil.copyfile(src, tmp)
            os.replace(tmp, dst)
            return True
        except (PermissionError, OSError) as e:
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                _dbg(f"copy 실패(최종): {src} -> {dst}: {type(e).__name__}: {e}")
    return False


def _safe_copytree(src: Path, dst: Path) -> None:
    """잠긴 파일을 재시도하며 디렉터리 복사"""
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        target = dst / item.name
        if item.is_dir():
            _safe_copytree(item, target)
        else:
            _copy_with_retry(item, target)


def _copy_bundle() -> None:
    app_dir = _app_dir()
    bundle = _bundle_dir()
    _dbg(f"=== _copy_bundle: app_dir={app_dir}, bundle={bundle} ===")
    for name in ("app.py", "config.py", "run_conversion.py",
                 "parsers", "converters", "classification",
                 "templates", "data"):
        src = bundle / name
        dst = app_dir / name
        if not src.exists():
            _dbg(f"  {name}: src 없음 ({src})")
            continue
        if src.is_dir():
            # 기존 디렉터리 삭제 (잠긴 파일은 무시)
            if dst.exists():
                shutil.rmtree(dst, ignore_errors=True)
            _safe_copytree(src, dst)
            _dbg(f"  {name}: 디렉터리 복사 완료 -> {dst}")
        else:
            try:
                shutil.copy2(src, dst)
                _dbg(f"  {name}: 파일 복사 완료 -> {dst}")
            except (PermissionError, OSError) as e:
                _dbg(f"  {name}: 파일 복사 실패: {type(e).__name__}: {e}")
                pass

    # PyInstaller가 한글 파일명 xlsx를 templates/ 하위가 아닌 번들 루트에
    # 배치하는 경우가 있으므로, 루트의 xlsx를 templates/로 추가 복사
    templates_dst = app_dir / "templates"
    templates_dst.mkdir(parents=True, exist_ok=True)
    for xlsx in bundle.glob("*.xlsx"):
        target = templates_dst / xlsx.name
        if not target.exists():
            try:
                shutil.copy2(xlsx, target)
            except (PermissionError, OSError):
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

    # 1.5) 같은 이름의 다른 인스턴스 종료 (templates/*.xlsx 잠금 방지)
    _kill_other_instances()

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
