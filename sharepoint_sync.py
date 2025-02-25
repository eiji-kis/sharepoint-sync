#!/usr/bin/env python3
"""
SharePoint synchronization CLI tool.

This script synchronizes files between two directories (e.g., KIS and Client SharePoints).
It supports:
  - Excluding specific files/directories
  - Detecting moved or updated files
  - Displaying diffs for Word documents
  - Interactive copy/move operations
  - Multiple profiles for different SharePoints

Requires:
  - python-docx
  - pyfiglet
  - click

Usage:
  - python sharepoint_sync.py setup
  - python sharepoint_sync.py sync <profile>
  - python sharepoint_sync.py exclude_dir <dir_name>
  - python sharepoint_sync.py exclude_file <file_name>
"""

from datetime import datetime
import os
import shutil
import filecmp
import json
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Set
import click
import difflib
import colorlog
import pyfiglet
from docx import Document

# ------------------------------------------------------------------------------
# Configuration Constants & Logger Setup
# ------------------------------------------------------------------------------
CONFIG_FILE = Path.home() / ".sharepoint_sync_profiles.json"
DEFAULT_EXCLUDED_FILES: List[str] = []
DEFAULT_EXCLUDED_DIRS: List[str] = []
FOLLOW_UP_FILE = Path("follow_up_tasks.md")

# Create a logger
logger = logging.getLogger("SharePoint Sync")
logger.setLevel(logging.DEBUG)  # Set base level; can override with CLI.
console_handler = logging.StreamHandler()
formatter = colorlog.ColoredFormatter(
        fmt="%(log_color)s[%(levelname)s]: %(message)s",
        datefmt=None,
        reset=True,
        log_colors={
            "DEBUG": "cyan",
            "INFO": "green",
            "WARNING": "yellow",
            "ERROR": "red",
            "CRITICAL": "bold_red",
        }
    )
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------
def log_follow_up(task: str) -> None:
    """
    Appends a task/instruction to 'follow_up_tasks.md', timestamped.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with FOLLOW_UP_FILE.open("a", encoding="utf-8") as f:
        f.write(f"### {timestamp}\n")
        f.write(task.strip() + "\n\n")

# ------------------------------------------------------------------------------
# Data Structures
# ------------------------------------------------------------------------------
class SyncProfile:
    """
    Holds the paths for a single synchronization profile (KIS and Client).
    """
    def __init__(self, kis_dir: str, client_dir: str) -> None:
        self.kis_dir = Path(kis_dir).resolve()
        self.client_dir = Path(client_dir).resolve()

    def validate(self) -> None:
        """
        Validates that KIS and Client directories exist.
        """
        if not self.kis_dir.exists():
            raise FileNotFoundError(f"KIS directory '{self.kis_dir}' does not exist.")
        if not self.client_dir.exists():
            raise FileNotFoundError(f"Client directory '{self.client_dir}' does not exist.")

# ------------------------------------------------------------------------------
# Configuration Management
# ------------------------------------------------------------------------------
def load_profiles() -> Dict:
    """
    Loads the profiles configuration file. Returns a dict with:
        {
            "excluded_files": [...],
            "excluded_dirs": [...],
            "profiles": {
                "profile_name": {
                    "kis_dir": "...",
                    "client_dir": "..."
                },
                ...
            }
        }
    """
    logger.debug(f"Loading profiles from {CONFIG_FILE}")
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as exc:
            logger.error(f"Failed to parse config file {CONFIG_FILE}: {exc}")
            # Return minimal default structure on error
            return {
                "excluded_files": DEFAULT_EXCLUDED_FILES,
                "excluded_dirs": DEFAULT_EXCLUDED_DIRS,
                "profiles": {}
            }
    else:
        # Return default structure if file doesn't exist
        return {
            "excluded_files": DEFAULT_EXCLUDED_FILES,
            "excluded_dirs": DEFAULT_EXCLUDED_DIRS,
            "profiles": {}
        }

def save_profiles(profiles: Dict) -> None:
    """
    Writes the profiles dictionary back to disk.
    """
    logger.debug(f"Saving profiles to {CONFIG_FILE}")
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(profiles, f, indent=4)
    except OSError as exc:
        logger.error(f"Failed to write config file {CONFIG_FILE}: {exc}")
        raise

# ------------------------------------------------------------------------------
# Word Document Handling
# ------------------------------------------------------------------------------
def extract_text_from_docx(file_path: Path) -> str:
    """
    Extracts text from a .docx file using the `docx` library.
    Returns the extracted text or an error message if something fails.
    """
    try:
        doc = Document(file_path)
        return "\n".join(para.text for para in doc.paragraphs)
    except Exception as exc:
        error_msg = f"[Error extracting text from {file_path}: {exc}]"
        logger.error(error_msg)
        return error_msg

def show_file_diff(file1: Path, file2: Path) -> None:
    """
    Displays a unified diff of two .docx files to the console.
    """
    try:
        file1_text = extract_text_from_docx(file1)
        file2_text = extract_text_from_docx(file2)

        if file1_text == file2_text:
            logger.info("The file content is the same on both SharePoints.")
            logger.info("Copying is recommended to synchronize the modified date.")
            return

        click.echo("\nShowing diff:")
        diff = difflib.unified_diff(
            file1_text.splitlines(),
            file2_text.splitlines(),
            fromfile=str(file1),
            tofile=str(file2),
            lineterm=''
        )
        for line in diff:
            if line.startswith("-"):
                click.echo(click.style(line, fg="red"))  # Removed content
            elif line.startswith("+"):
                click.echo(click.style(line, fg="green"))  # Added content
            else:
                click.echo(line)
        click.echo("")
    except Exception as exc:
        logger.exception(f"Could not generate diff for {file1} vs {file2}: {exc}")

# ------------------------------------------------------------------------------
# SharePoint Comparison
# ------------------------------------------------------------------------------
def compare_sharepoints(
    kis_dir: Path,
    client_dir: Path,
    excluded_files: List[str],
    excluded_dirs: List[str]
) -> Tuple[Set[str], Set[str], Dict[str, Tuple[str, str]], Dict[str, Path]]:
    """
    Compare two directories (kis_dir and client_dir), ignoring the given excluded
    files and directories. Returns a tuple of:
      (kis_only, client_only, moved_files, updated_files)

    kis_only: Set of relative paths that exist only in KIS.
    client_only: Set of relative paths that exist only in the client.
    moved_files: Dict[file_name -> (kis_relative_path, client_relative_path)]
    updated_files: Dict[relative_path -> Path_of_latest_file]
    """
    logger.debug("Comparing KIS and Client SharePoints...")
    logger.debug(f"Excluded files: {excluded_files}")
    logger.debug(f"Excluded dirs: {excluded_dirs}")

    def should_exclude(file_path: Path) -> bool:
        if file_path.name in excluded_files:
            logger.debug(f"Excluding file due to name: {file_path}")
            return True
        if any(excl in file_path.parts for excl in excluded_dirs):
            logger.debug(f"Excluding file due to directory path: {file_path}")
            return True
        return False

    kis_files = {
        str(file.relative_to(kis_dir)): file
        for file in kis_dir.rglob("*")
        if file.is_file() and not should_exclude(file)
    }
    client_files = {
        str(file.relative_to(client_dir)): file
        for file in client_dir.rglob("*")
        if file.is_file() and not should_exclude(file)
    }

    kis_only = set(kis_files) - set(client_files)
    client_only = set(client_files) - set(kis_files)
    common_files = set(kis_files) & set(client_files)
    moved_files = {}
    updated_files = {}

    # Detect moved files by matching filenames in different relative paths
    for file_rel_path in kis_only.copy():
        file_name = Path(file_rel_path).name
        matching_file = next(
            (cf for cf in client_only if Path(cf).name == file_name),
            None
        )
        if matching_file:
            moved_files[file_name] = (file_rel_path, matching_file)
            kis_only.remove(file_rel_path)
            client_only.remove(matching_file)

    # Detect updated files in the common set
    for file_rel_path in common_files:
        kis_file = kis_files[file_rel_path]
        client_file = client_files[file_rel_path]
        if kis_file.stat().st_mtime > client_file.stat().st_mtime:
            updated_files[file_rel_path] = kis_file
        elif client_file.stat().st_mtime > kis_file.stat().st_mtime:
            updated_files[file_rel_path] = client_file

    return kis_only, client_only, moved_files, updated_files

# ------------------------------------------------------------------------------
# CLI Utility Logging
# ------------------------------------------------------------------------------
def log_separator(char: str = "-", length: int = 50) -> None:
    """
    Prints a colored separator line to the console.
    """
    click.echo(click.style(char * length, fg="cyan"))

def log_ascii_message(message: str = "SharePoint Sync", font: str = "slant") -> None:
    """
    Logs a message in ASCII art using pyfiglet.
    """
    ascii_art = pyfiglet.figlet_format(message, font=font)
    click.echo(click.style(ascii_art, fg="magenta"))

# ------------------------------------------------------------------------------
# CLI Commands
# ------------------------------------------------------------------------------
@click.group()
@click.option(
    "--verbosity",
    default="info",
    type=click.Choice(["debug", "info", "warning", "error", "critical"], case_sensitive=False),
    help="Set the logging verbosity level."
)
def cli(verbosity: str) -> None:
    """
    Main entry point for the SharePoint Sync CLI.
    """
    log_level = getattr(logging, verbosity.upper(), logging.INFO)
    logger.setLevel(log_level)
    logger.debug(f"Logger set to {verbosity.upper()} level.")

@cli.command()
@click.argument("profile", required=True)
def sync(profile: str) -> None:
    """
    Synchronize the specified PROFILE between KIS and Client.
    """
    log_ascii_message("SharePoint Sync")

    profiles = load_profiles()
    if profile not in profiles["profiles"]:
        logger.warning(f"Profile '{profile}' not found. Run 'setup' command first.")
        return

    # Load profile data
    try:
        profile_data = profiles["profiles"][profile]
        sync_profile = SyncProfile(
            kis_dir=profile_data["kis_dir"],
            client_dir=profile_data["client_dir"]
        )
        sync_profile.validate()
        excluded_files = profiles.get("excluded_files", [])
        excluded_dirs = profiles.get("excluded_dirs", [])
    except KeyError as exc:
        logger.error(f"Missing required key in profile data: {exc}")
        return
    except FileNotFoundError as exc:
        logger.error(f"{exc}")
        return

    # Compare sharepoints
    kis_only, client_only, moved_files, updated_files = compare_sharepoints(
        sync_profile.kis_dir,
        sync_profile.client_dir,
        excluded_files,
        excluded_dirs
    )

    to_be_created_client = len(kis_only)
    to_be_created_kis = len(client_only)
    to_be_moved = len(moved_files)
    to_be_updated = len(updated_files)

    log_separator()
    logger.info("Sync Summary:")
    logger.info(f"- {to_be_created_client} files to be created on {profile}.")
    logger.info(f"- {to_be_created_kis} files to be created on KIS.")
    logger.info(f"- {to_be_moved} files have been moved.")
    logger.info(f"- {to_be_updated} files have been updated.")
    log_separator()

    if not any([to_be_created_client, to_be_created_kis, to_be_moved, to_be_updated]):
        logger.info(f"KIS and {profile} SharePoints are already in sync. 😎")
        return

    # Handle files that exist only in KIS -> create on Client
    if to_be_created_client:
        logger.debug("Handling files that exist only on KIS.")
        for relative_path in kis_only:
            logger.info(f"File {relative_path} found on KIS only, missing on {profile}.")
            origin = sync_profile.kis_dir / relative_path
            target = sync_profile.client_dir / relative_path
            if click.confirm(f"Copy {relative_path} from KIS to {profile}?"):
                os.makedirs(target.parent, exist_ok=True)
                shutil.copy2(origin, target)
            else:
                log_follow_up(
                    f"You chose NOT to copy '{relative_path}' from KIS to {profile}.\n"
                    f"Please manually check:\n"
                    f" - KIS path: {origin}\n"
                    f" - {profile} path: {target}\n"
                )

    # Handle files that exist only in Client -> create on KIS
    if to_be_created_kis:
        logger.debug("Handling files that exist only on Client.")
        for relative_path in client_only:
            logger.info(f"File {relative_path} found on {profile} only, missing on KIS.")
            origin = sync_profile.client_dir / relative_path
            target = sync_profile.kis_dir / relative_path
            if click.confirm(f"Copy {relative_path} from {profile} to KIS?"):
                os.makedirs(target.parent, exist_ok=True)
                shutil.copy2(origin, target)
            else:
                # User chose NO, so log a follow-up task
                log_follow_up(
                    f"You chose NOT to copy '{relative_path}' from {profile} to KIS.\n"
                    f"Please manually check:\n"
                    f" - {profile} path: {origin}\n"
                    f" - KIS path: {target}\n"
                )

    # Handle moved files
    if to_be_moved:
        logger.debug("Handling moved files.")
        for file_name, (kis_rel_path, client_rel_path) in moved_files.items():
            kis_abs_path = sync_profile.kis_dir / kis_rel_path
            client_abs_path = sync_profile.client_dir / client_rel_path

            logger.debug(f"KIS Relative Path: {kis_rel_path}")
            logger.debug(f"Client Relative Path: {client_rel_path}")

            if (
                kis_abs_path.exists()
                and client_abs_path.exists()
            ):
                
                # Choose the "latest" by creation time to keep the directory structure
                if kis_abs_path.stat().st_ctime > client_abs_path.stat().st_ctime:
                    latest_abs = kis_abs_path
                    outdated_abs = client_abs_path
                    latest_sharepoint = "KIS"
                    outdated_sharepoint = profile
                    outdated_rel_path = client_rel_path
                    latest_rel_path = kis_rel_path
                else:
                    latest_abs = client_abs_path
                    outdated_abs = kis_abs_path
                    latest_sharepoint = profile
                    outdated_sharepoint = "KIS"
                    outdated_rel_path = kis_rel_path
                    latest_rel_path = client_rel_path

                if (filecmp.cmp(kis_abs_path, client_abs_path, shallow=False)):
                    logger.debug("File exists in both locations and they are identical.")
                    logger.info(
                    f"The file '{file_name}' was moved on {latest_sharepoint} "
                    f"(file content is the same in both)."
                    )
                else:
                    # If they are not identical, that suggests a conflict or partial move scenario
                    logger.warning(
                        f"Potentially conflicting move for '{file_name}'. "
                        "Files differ and were moved. Proceed with caution."
                    )
                    if latest_rel_path.endswith(".docx"):
                        show_file_diff(outdated_abs, latest_abs)

                
                destination_dir = (
                    sync_profile.client_dir / latest_rel_path
                ).parent if outdated_sharepoint == profile else (
                    sync_profile.kis_dir / latest_rel_path
                ).parent

                source_path = (
                    sync_profile.client_dir / outdated_rel_path
                ) if outdated_sharepoint == profile else (
                    sync_profile.kis_dir / outdated_rel_path
                )

                destination_path = (
                    sync_profile.client_dir / latest_rel_path
                ) if outdated_sharepoint == profile else (
                    sync_profile.kis_dir / latest_rel_path
                )        

                if click.confirm(
                    f"Move {file_name} on {outdated_sharepoint} from '/{outdated_rel_path}' "
                    f"to '/{latest_rel_path}' to match {latest_sharepoint}?"
                ):

                    try:
                        os.makedirs(destination_dir, exist_ok=True)
                        shutil.move(str(source_path), str(destination_path))
                        logger.info(f"Successfully moved {file_name} to match structure.")
                    except FileNotFoundError:
                        logger.error(f"Move failed: {source_path} does not exist!")
                    except PermissionError:
                        logger.error(f"Move failed: Permission denied.")
                    except shutil.Error as err:
                        logger.error(f"Move failed: {err}")
                else:
                    log_follow_up(
                        f"You chose NOT to move '{file_name}' from '{outdated_sharepoint}' "
                        f"path '/{outdated_rel_path}' to '/{latest_rel_path}' to match '{latest_sharepoint}'.\n"
                        f"Please manually check:\n"
                        f" - Outdated file path: {source_path}\n"
                        f" - Destination path: {destination_path}\n"
                    )
            

    # Handle updated files
    if to_be_updated:
        logger.debug("Handling updated files.")
        for rel_path, latest_file_abs in updated_files.items():
            kis_abs_path = sync_profile.kis_dir / rel_path
            client_abs_path = sync_profile.client_dir / rel_path

            if latest_file_abs == kis_abs_path:
                latest_sharepoint = "KIS"
                outdated_sharepoint = profile
                outdated_file_abs = client_abs_path
            else:
                latest_sharepoint = profile
                outdated_sharepoint = "KIS"
                outdated_file_abs = kis_abs_path

            logger.info(
                f"The file '{rel_path}' was modified on {latest_sharepoint}."
            )

            # Display diff if it's a docx file
            if rel_path.endswith(".docx"):
                show_file_diff(outdated_file_abs, latest_file_abs)

            if click.confirm(
                f"Copy '{rel_path}' from {latest_sharepoint} to {outdated_sharepoint}?"
            ):
                try:
                    shutil.copy2(latest_file_abs, outdated_file_abs)
                    logger.info(
                        f"Successfully copied '{rel_path}' "
                        f"from {latest_sharepoint} to {outdated_sharepoint}."
                    )
                except FileNotFoundError:
                    logger.error(f"Source file {latest_file_abs} does not exist.")
                except PermissionError:
                    logger.error(f"Permission denied while copying {latest_file_abs}.")
                except shutil.SameFileError:
                    logger.warning("Source and destination are the same file.")
                except Exception as exc:
                    logger.exception(f"Unexpected error during file copy: {exc}")
            else:
                log_follow_up(
                    f"You chose NOT to copy '{rel_path}' from {latest_sharepoint} to {outdated_sharepoint}.\n"
                    f"Please manually check:\n"
                    f" - Latest file: {latest_file_abs}\n"
                    f" - Outdated file: {outdated_file_abs}\n"
                )

@cli.command()
@click.argument("dir_name", required=True)
def exclude_dir(dir_name: str) -> None:
    """
    Exclude a directory (by name or partial path) from sync operations.
    """
    profiles = load_profiles()
    if dir_name not in profiles["excluded_dirs"]:
        profiles["excluded_dirs"].append(dir_name)
        save_profiles(profiles)
        click.echo(f"✅ Directory '{dir_name}' added to exclusion list.")
    else:
        click.echo(f"Directory '{dir_name}' is already excluded.")

@cli.command()
@click.argument("file_name", required=True)
def exclude_file(file_name: str) -> None:
    """
    Exclude a file (by name) from sync operations.
    """
    profiles = load_profiles()
    if file_name not in profiles["excluded_files"]:
        profiles["excluded_files"].append(file_name)
        save_profiles(profiles)
        click.echo(f"✅ File '{file_name}' added to exclusion list.")
    else:
        click.echo(f"File '{file_name}' is already excluded.")

@cli.command()
@click.option("--name", prompt="Profile name (e.g. client name)", help="Name of the profile")
@click.option("--kis-dir", prompt="Root directory of KIS SharePoint", type=click.Path(exists=True))
@click.option("--client-dir", prompt="Root directory of Client SharePoint", type=click.Path(exists=True))
def setup(name: str, kis_dir: str, client_dir: str) -> None:
    """
    Setup a new sync profile by providing directories for KIS and Client SharePoints.
    """
    profiles = load_profiles()
    profiles["profiles"][name] = {"kis_dir": kis_dir, "client_dir": client_dir}
    save_profiles(profiles)
    click.echo(f"Profile '{name}' saved successfully.")

# ------------------------------------------------------------------------------
# Main Entry
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    cli()
