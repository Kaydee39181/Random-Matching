"""Shared installer, launcher, expiry, cleanup, and task utilities.

This module is intentionally self-contained and uses only the Python standard
library so it can run on clean Windows machines with Python 3.12+ installed.
"""

from __future__ import annotations

import argparse
import ctypes
import hmac
import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Sequence

import config


CHECKSUM_FIELD = "checksum"
CHECKSUM_VERSION = "sha256-v1"


class RandomMatchingError(RuntimeError):
    """Base exception for managed installer and launcher failures."""


class IntegrityError(RandomMatchingError):
    """Raised when installation metadata fails integrity checks."""


class Logger:
    """Factory for rotating file loggers used by every component."""

    @staticmethod
    def configure(name: str, log_file: Path, level: int = logging.INFO) -> logging.Logger:
        """Create a logger that writes to a rotating log file and stderr."""
        log_file.parent.mkdir(parents=True, exist_ok=True)
        logger = logging.getLogger(name)
        logger.setLevel(level)
        logger.propagate = False

        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            handler.close()

        formatter = logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=1_000_000,
            backupCount=5,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        logger.addHandler(stream_handler)
        return logger


@dataclass(frozen=True)
class InstallMetadata:
    """Persisted metadata used for expiry and basic tamper checks."""

    install_date: str
    expiry_date: str
    software_version: str
    installation_path: str
    installation_uuid: str
    checksum_version: str = CHECKSUM_VERSION
    checksum: str = ""

    @classmethod
    def create(cls, installation_path: Path, expiry_days: int) -> "InstallMetadata":
        """Create unsigned metadata for a new or refreshed installation."""
        now = datetime.now().replace(microsecond=0)
        expiry = now + timedelta(days=expiry_days)
        return cls(
            install_date=now.isoformat(),
            expiry_date=expiry.isoformat(),
            software_version=config.SOFTWARE_VERSION,
            installation_path=str(installation_path.resolve()),
            installation_uuid=str(uuid.uuid4()),
        )

    @classmethod
    def from_file(cls, path: Path) -> "InstallMetadata":
        """Read metadata from disk."""
        data = json.loads(path.read_text(encoding="utf-8"))
        return cls(**data)

    def to_dict(self, include_checksum: bool = True) -> dict[str, str]:
        """Convert metadata to a JSON-serializable dictionary."""
        data = {
            "install_date": self.install_date,
            "expiry_date": self.expiry_date,
            "software_version": self.software_version,
            "installation_path": self.installation_path,
            "installation_uuid": self.installation_uuid,
            "checksum_version": self.checksum_version,
        }
        if include_checksum:
            data[CHECKSUM_FIELD] = self.checksum
        return data

    def signed(self) -> "InstallMetadata":
        """Return a copy with its checksum populated."""
        checksum = compute_metadata_checksum(self.to_dict(include_checksum=False))
        return InstallMetadata(**self.to_dict(include_checksum=False), checksum=checksum)

    def verify_checksum(self) -> None:
        """Validate the metadata checksum."""
        expected = compute_metadata_checksum(self.to_dict(include_checksum=False))
        if not self.checksum or not hmac_compare(self.checksum, expected):
            raise IntegrityError("install.json checksum verification failed")

    @property
    def expiry_datetime(self) -> datetime:
        """Return the parsed expiry datetime."""
        return datetime.fromisoformat(self.expiry_date)

    @property
    def installation_dir(self) -> Path:
        """Return the installation directory path."""
        return Path(self.installation_path)


def compute_metadata_checksum(data: dict[str, str]) -> str:
    """Compute a deterministic checksum over canonical metadata JSON."""
    canonical = json.dumps(data, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def hmac_compare(left: str, right: str) -> bool:
    """Compare two checksum strings without leaking length-normal timing."""
    return hmac.compare_digest(left, right)


def quote_cmd(value: str) -> str:
    """Quote a string for use in a generated Windows batch file."""
    return '"' + value.replace('"', '""') + '"'


class CleanupManager:
    """Generate and invoke the ProgramData cleanup batch file."""

    def __init__(
        self,
        install_path: Path,
        logger: logging.Logger,
        dry_run: bool = config.ENABLE_DRY_RUN,
    ) -> None:
        self.install_path = install_path.resolve()
        self.logger = logger
        self.dry_run = dry_run

    def write_cleanup_script(self) -> Path:
        """Generate cleanup.bat outside the application folder."""
        config.PROGRAMDATA_PATH.mkdir(parents=True, exist_ok=True)
        content = self.render_cleanup_script()
        config.CLEANUP_FILE.write_text(content, encoding="utf-8", newline="\r\n")
        self.verify_cleanup_script()
        self.logger.info("Cleanup script written: %s", config.CLEANUP_FILE)
        return config.CLEANUP_FILE

    def render_cleanup_script(self) -> str:
        """Render the cleanup batch script with the current configuration."""
        standard_processes = ("python.exe", "pythonw.exe", "node.exe")
        all_processes = list(dict.fromkeys([*standard_processes, *config.CUSTOM_PROCESS_NAMES]))
        taskkill_lines = "\n".join(
            f"taskkill /F /IM {quote_cmd(process_name)} /T >nul 2>nul"
            for process_name in all_processes
        )
        delete_lines = (
            f'echo DRY RUN: would delete {quote_cmd(str(self.install_path))} >> "%LOG_FILE%"\n'
            f'echo DRY RUN: would delete scheduled task {config.TASK_NAME} >> "%LOG_FILE%"'
            if self.dry_run
            else (
                f'if exist {quote_cmd(str(self.install_path))} rmdir /S /Q {quote_cmd(str(self.install_path))}\n'
                f'schtasks /Delete /TN {quote_cmd(config.TASK_NAME)} /F >nul 2>nul\n'
                f'if exist {quote_cmd(str(config.INSTALL_FILE))} del /F /Q {quote_cmd(str(config.INSTALL_FILE))} >nul 2>nul\n'
                f'if exist {quote_cmd(str(config.TASK_XML_FILE))} del /F /Q {quote_cmd(str(config.TASK_XML_FILE))} >nul 2>nul\n'
                f'if exist {quote_cmd(str(config.PROGRAMDATA_PATH / "install.log"))} del /F /Q {quote_cmd(str(config.PROGRAMDATA_PATH / "install.log"))} >nul 2>nul'
            )
        )
        delayed_self_delete = (
            "start \"\" /min cmd /c \"ping 127.0.0.1 -n 3 >nul & del /F /Q \"\"%~f0\"\" >nul 2>nul\""
            if not self.dry_run
            else 'echo DRY RUN: would delete cleanup.bat >> "%LOG_FILE%"'
        )
        return f"""@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "LOG_FILE={config.PROGRAMDATA_PATH}\\cleanup.log"
echo [%DATE% %TIME%] Cleanup started. >> "%LOG_FILE%"
timeout /T 5 /NOBREAK >nul 2>nul

rem Terminate common runtimes and editable custom executable names.
{taskkill_lines}

rem Delete only the Random Matching installation folder, never ProgramData.
{delete_lines}

echo [%DATE% %TIME%] Cleanup finished. >> "%LOG_FILE%"
{delayed_self_delete}
exit /b 0
"""

    def verify_cleanup_script(self) -> None:
        """Check that cleanup.bat exists and targets the application folder."""
        if not config.CLEANUP_FILE.exists():
            raise RandomMatchingError("cleanup.bat was not created")
        content = config.CLEANUP_FILE.read_text(encoding="utf-8")
        if str(self.install_path) not in content:
            raise RandomMatchingError("cleanup.bat does not reference the installation path")
        if str(config.PROGRAMDATA_PATH) in str(self.install_path):
            raise RandomMatchingError("installation path cannot be inside ProgramData")

    def invoke_cleanup(self, silent: bool = True) -> None:
        """Start cleanup.bat without waiting for it to finish."""
        if not config.CLEANUP_FILE.exists():
            self.write_cleanup_script()
        creation_flags = subprocess.CREATE_NO_WINDOW if silent and os.name == "nt" else 0
        subprocess.Popen(
            ["cmd.exe", "/c", str(config.CLEANUP_FILE)],
            cwd=str(config.PROGRAMDATA_PATH),
            creationflags=creation_flags,
        )
        self.logger.info("Cleanup invoked: %s", config.CLEANUP_FILE)


class TaskScheduler:
    """Create, delete, and verify the Windows scheduled expiry task."""

    def __init__(self, logger: logging.Logger) -> None:
        self.logger = logger

    def exists(self) -> bool:
        """Return True if the scheduled task exists."""
        result = subprocess.run(
            ["schtasks", "/Query", "/TN", config.TASK_NAME],
            capture_output=True,
            text=True,
            check=False,
        )
        return result.returncode == 0

    def create_or_update(self, run_at: datetime, cleanup_file: Path) -> None:
        """Create or update the expiry scheduled task from XML."""
        if self.exists():
            self.delete()
        xml_path = self.write_task_xml(run_at, cleanup_file)
        result = subprocess.run(
            ["schtasks", "/Create", "/TN", config.TASK_NAME, "/XML", str(xml_path), "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            raise RandomMatchingError(f"scheduled task creation failed: {result.stderr.strip()}")
        if not self.exists():
            raise RandomMatchingError("scheduled task verification failed after creation")
        self.logger.info("Scheduled task created or updated: %s", config.TASK_NAME)

    def write_task_xml(self, run_at: datetime, cleanup_file: Path) -> Path:
        """Write the task definition XML with StartWhenAvailable enabled."""
        config.PROGRAMDATA_PATH.mkdir(parents=True, exist_ok=True)
        start_boundary = run_at.replace(microsecond=0).isoformat()
        xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Expires and removes the Random Matching evaluation copy.</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>cmd.exe</Command>
      <Arguments>/c "{cleanup_file}"</Arguments>
      <WorkingDirectory>{config.PROGRAMDATA_PATH}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""
        config.TASK_XML_FILE.write_bytes(xml.encode("utf-16"))
        self.logger.info("Task XML written: %s", config.TASK_XML_FILE)
        return config.TASK_XML_FILE

    def verify(self) -> None:
        """Raise if the scheduled task is missing."""
        if not self.exists():
            raise IntegrityError("scheduled task is missing")

    def delete(self) -> None:
        """Delete the scheduled task if present."""
        result = subprocess.run(
            ["schtasks", "/Delete", "/TN", config.TASK_NAME, "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode == 0:
            self.logger.info("Scheduled task deleted: %s", config.TASK_NAME)
        else:
            self.logger.info("Scheduled task delete skipped or failed: %s", result.stderr.strip())


class ExpiryManager:
    """Read metadata and enforce expiry/tamper policy."""

    def __init__(self, logger: logging.Logger) -> None:
        self.logger = logger

    def load_metadata(self) -> InstallMetadata:
        """Load and validate install.json."""
        if not config.INSTALL_FILE.exists():
            raise IntegrityError("install.json is missing")
        metadata = InstallMetadata.from_file(config.INSTALL_FILE)
        metadata.verify_checksum()
        return metadata

    def verify(self, current_install_path: Path) -> InstallMetadata:
        """Verify integrity, path, cleanup file, task, and expiry."""
        metadata = self.load_metadata()
        stored_path = metadata.installation_dir.resolve()
        current_path = current_install_path.resolve()
        if stored_path != current_path:
            raise IntegrityError(f"installation folder moved from {stored_path} to {current_path}")
        if config.ENABLE_TAMPER_CHECK:
            if not config.CLEANUP_FILE.exists():
                raise IntegrityError("cleanup.bat is missing")
            TaskScheduler(self.logger).verify()
        if datetime.now() >= metadata.expiry_datetime:
            raise RandomMatchingError("evaluation period expired")
        self.logger.info("Expiry verification passed. Expires: %s", metadata.expiry_date)
        return metadata

    def write_metadata(self, metadata: InstallMetadata) -> None:
        """Write signed installation metadata."""
        config.PROGRAMDATA_PATH.mkdir(parents=True, exist_ok=True)
        signed = metadata.signed()
        config.INSTALL_FILE.write_text(
            json.dumps(signed.to_dict(), indent=2, sort_keys=True),
            encoding="utf-8",
        )
        self.logger.info("Installation metadata written: %s", config.INSTALL_FILE)


class Installer:
    """Install or refresh expiry enforcement for the bundle."""

    def __init__(self, bundle_root: Path | None = None) -> None:
        self.bundle_root = (bundle_root or Path(__file__).resolve().parent).resolve()
        self.logger = Logger.configure("installer", config.PROGRAMDATA_PATH / "install.log")
        self.created_files: list[Path] = []

    def install(self) -> int:
        """Run the full installation and rollback on failure."""
        self.logger.info("Install started")
        self.logger.info("Detected folder: %s", self.bundle_root)
        metadata = InstallMetadata.create(self.bundle_root, config.EXPIRY_DAYS)
        cleanup = CleanupManager(self.bundle_root, self.logger)
        scheduler = TaskScheduler(self.logger)
        expiry = ExpiryManager(self.logger)

        try:
            config.PROGRAMDATA_PATH.mkdir(parents=True, exist_ok=True)
            config.LOG_DIRECTORY.mkdir(parents=True, exist_ok=True)
            self.logger.info("Install time: %s", metadata.install_date)
            self.logger.info("Expiry date: %s", metadata.expiry_date)

            expiry.write_metadata(metadata)
            self.created_files.append(config.INSTALL_FILE)

            cleanup_path = cleanup.write_cleanup_script()
            self.created_files.append(cleanup_path)
            self.logger.info("Cleanup script location: %s", cleanup_path)

            scheduler.create_or_update(metadata.expiry_datetime, cleanup_path)
            self.logger.info("Task creation status: success")

            self.verify_generated_project_files()
            expiry.verify(self.bundle_root)
            self.logger.info("Install completed successfully")
            return 0
        except Exception as exc:
            self.logger.exception("Install failed: %s", exc)
            self.rollback(scheduler)
            return 1

    def verify_generated_project_files(self) -> None:
        """Ensure required root modules are present for deployment."""
        for file_name in ("config.py", "random_matching_core.py", "launcher.py", "uninstall.py"):
            path = self.bundle_root / file_name
            if not path.exists():
                raise RandomMatchingError(f"required file is missing: {path}")

    def rollback(self, scheduler: TaskScheduler) -> None:
        """Remove partially created files and scheduled task after failure."""
        self.logger.info("Rollback started")
        try:
            scheduler.delete()
        except Exception as exc:
            self.logger.error("Rollback could not delete scheduled task: %s", exc)
        for path in reversed(self.created_files):
            try:
                if path.exists():
                    path.unlink()
                    self.logger.info("Rollback removed: %s", path)
            except Exception as exc:
                self.logger.error("Rollback could not remove %s: %s", path, exc)
        self.logger.info("Rollback finished")


class Launcher:
    """Verify expiry state, then launch one configured application."""

    def __init__(self, bundle_root: Path | None = None) -> None:
        self.bundle_root = (bundle_root or Path(__file__).resolve().parent).resolve()
        self.logger = Logger.configure("launcher", config.LOG_DIRECTORY / "launcher.log")

    def run(self, app_name: str | None, passthrough: Sequence[str] | None = None) -> int:
        """Verify the installation and launch the requested application."""
        try:
            self.verify_or_repair()
            command, cwd = self.resolve_command(app_name, passthrough or [])
            self.logger.info("Launching app. cwd=%s command=%s", cwd, command)
            return subprocess.call(command, cwd=str(cwd))
        except RandomMatchingError as exc:
            self.logger.error("Launch refused: %s", exc)
            self.handle_refusal(str(exc))
            return 20
        except Exception as exc:
            self.logger.exception("Unexpected launch failure: %s", exc)
            self.handle_refusal("The application could not start because of an unexpected error.")
            return 21

    def verify_or_repair(self) -> None:
        """Perform integrity checks and repair cleanup/task when possible."""
        expiry = ExpiryManager(self.logger)
        cleanup = CleanupManager(self.bundle_root, self.logger)
        scheduler = TaskScheduler(self.logger)
        try:
            metadata = expiry.verify(self.bundle_root)
            self.logger.info("Launch verification succeeded for UUID %s", metadata.installation_uuid)
        except IntegrityError as exc:
            self.logger.warning("Tamper or integrity issue detected: %s", exc)
            if not config.CLEANUP_FILE.exists():
                cleanup.write_cleanup_script()
                self.logger.info("Recreated cleanup script after tamper detection")
            try:
                metadata = expiry.load_metadata()
                scheduler.create_or_update(metadata.expiry_datetime, config.CLEANUP_FILE)
                self.logger.info("Recreated scheduled task after tamper detection")
            except Exception as repair_exc:
                self.logger.error("Repair attempt failed: %s", repair_exc)
            if config.BEGIN_CLEANUP_ON_TAMPER:
                cleanup.invoke_cleanup()
            raise
        except RandomMatchingError as exc:
            self.logger.info("Expiry reached or policy failure: %s", exc)
            self.show_expired_message()
            cleanup.invoke_cleanup()
            raise

    def resolve_command(
        self,
        app_name: str | None,
        passthrough: Sequence[str],
    ) -> tuple[list[str], Path]:
        """Resolve a launcher alias or explicit command."""
        if passthrough:
            return list(passthrough), self.bundle_root
        if not app_name:
            aliases = ", ".join(sorted(config.APPLICATIONS))
            raise RandomMatchingError(f"no application selected; available aliases: {aliases}")
        app = config.APPLICATIONS.get(app_name)
        if app is None:
            raise RandomMatchingError(f"unknown application alias: {app_name}")
        cwd = self.bundle_root / str(app["cwd"])
        command = [str(part) for part in app["command"]]  # type: ignore[index]
        if not cwd.exists():
            raise RandomMatchingError(f"application folder is missing: {cwd}")
        return command, cwd

    def handle_refusal(self, detail: str) -> None:
        """Show a professional refusal message."""
        if "expired" in detail.lower():
            self.show_expired_message()
        else:
            show_message(
                "Random Matching",
                "This evaluation copy cannot be launched.\nPlease contact the developer for support.",
            )

    def show_expired_message(self) -> None:
        """Display the standard expiry message."""
        show_message(
            "Evaluation expired",
            "This evaluation copy has expired.\nPlease contact the developer for a renewed license.",
        )


class Uninstaller:
    """Remove task, metadata, ProgramData files, and optionally app files."""

    def __init__(self, bundle_root: Path | None = None) -> None:
        self.bundle_root = (bundle_root or Path(__file__).resolve().parent).resolve()
        self.logger = Logger.configure("uninstaller", config.LOG_DIRECTORY / "uninstall.log")

    def uninstall(self, remove_application: bool = False) -> int:
        """Run uninstallation and return a process exit code."""
        try:
            self.logger.info("Uninstall started")
            self.terminate_processes()
            TaskScheduler(self.logger).delete()
            if remove_application:
                CleanupManager(self.bundle_root, self.logger).write_cleanup_script()
            self.remove_programdata_files()
            if remove_application:
                CleanupManager(self.bundle_root, self.logger).invoke_cleanup()
            self.logger.info("Uninstall completed")
            return 0
        except Exception as exc:
            self.logger.exception("Uninstall failed: %s", exc)
            return 1

    def terminate_processes(self) -> None:
        """Terminate runtimes that commonly lock application files."""
        current_pid = str(os.getpid())
        for process_name in ("python.exe", "pythonw.exe", "node.exe", *config.CUSTOM_PROCESS_NAMES):
            result = subprocess.run(
                [
                    "taskkill",
                    "/F",
                    "/FI",
                    f"IMAGENAME eq {process_name}",
                    "/FI",
                    f"PID ne {current_pid}",
                    "/T",
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            self.logger.info(
                "Terminate attempted for process: %s status=%s",
                process_name,
                result.returncode,
            )

    def remove_programdata_files(self) -> None:
        """Delete metadata, task XML, cleanup script, and logs in ProgramData."""
        for path in (config.INSTALL_FILE, config.TASK_XML_FILE):
            if path.exists():
                path.unlink()
                self.logger.info("Removed file: %s", path)
        if config.CLEANUP_FILE.exists():
            config.CLEANUP_FILE.unlink()
            self.logger.info("Removed file: %s", config.CLEANUP_FILE)

        self.close_logger_handlers()
        if config.LOG_DIRECTORY.exists():
            shutil.rmtree(config.LOG_DIRECTORY, ignore_errors=True)
        install_log = config.PROGRAMDATA_PATH / "install.log"
        if install_log.exists():
            install_log.unlink()
        try:
            config.PROGRAMDATA_PATH.rmdir()
        except OSError:
            pass

    def close_logger_handlers(self) -> None:
        """Release log file handles before deleting log directories on Windows."""
        for handler in list(self.logger.handlers):
            self.logger.removeHandler(handler)
            handler.close()


def show_message(title: str, message: str) -> None:
    """Show a Windows message box, falling back to console output."""
    try:
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x00000010)
    except Exception:
        print(f"{title}: {message}", file=sys.stderr)


def build_launcher_parser() -> argparse.ArgumentParser:
    """Build CLI parser for launcher.py."""
    parser = argparse.ArgumentParser(description="Random Matching guarded launcher")
    parser.add_argument("app", nargs="?", help="Configured application alias to launch")
    parser.add_argument(
        "--cmd",
        nargs=argparse.REMAINDER,
        help="Explicit command to run after verification, for advanced integrations",
    )
    return parser
