import os
import json
import shutil
import filecmp
import click
import difflib
from pathlib import Path
from datetime import datetime
from docx import Document

CONFIG_FILE = Path.home() / ".sharepoint_sync_profiles.json"

# Load configuration with default structure
def load_profiles():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"excluded_files": [], "excluded_dirs": [], "profiles": {}}

# Save profiles with updated structure
def save_profiles(profiles):
    with open(CONFIG_FILE, "w") as f:
        json.dump(profiles, f, indent=4)

def extract_text_from_docx(file_path):
    "Extract text from a .docx file."
    try:
        doc = Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        return f"[Error extracting text: {e}]"

def show_file_diff(file1, file2, source_name):
    "Displays a colorized diff of two files, handling both text and Word documents."
    try:
        if file1.suffix == ".docx" and file2.suffix == ".docx":
            file1_text = extract_text_from_docx(file1)
            file2_text = extract_text_from_docx(file2)
        else:
            click.echo("Unable to show diff: Only .docx files are supported.")
            return
        
        if source_name == 'KIS':
            fromfile, tofile = str(file2), str(file1)  
            file1_lines, file2_lines = file2_text.splitlines(), file1_text.splitlines()
        else:
            fromfile, tofile = str(file1), str(file2)  
            file1_lines, file2_lines = file1_text.splitlines(), file2_text.splitlines()
    
        diff = difflib.unified_diff(file1_lines, file2_lines, fromfile=fromfile, tofile=tofile)
        
        for line in diff:
            if line.startswith("-"):
                click.echo(click.style(line, fg="red"))  # Red for removed content
            elif line.startswith("+"):
                click.echo(click.style(line, fg="green"))  # Green for added content
            else:
                click.echo(line)
    
    except Exception as e:
        click.echo(f"Could not generate diff: {e}")

@click.group()
def cli():
    "CLI tool to manually sync SharePoint directories"
    pass

@click.command()
@click.option("--name", prompt="Profile name (e.g., client name)", help="Name of the profile")
@click.option("--kis-dir", prompt="Root directory of KIS SharePoint", type=click.Path(exists=True))
@click.option("--client-dir", prompt="Root directory of Client SharePoint", type=click.Path(exists=True))
def setup(name, kis_dir, client_dir):
    "Setup a new SharePoint sync profile"
    profiles = load_profiles()
    profiles["profiles"][name] = {"kis_dir": kis_dir, "client_dir": client_dir}
    save_profiles(profiles)
    click.echo(f"Profile '{name}' saved successfully.")

@click.command()
@click.argument("profile")
def sync(profile):
    "Sync files between KIS and {profile} SharePoint for a given profile"
    profiles = load_profiles()
    if profile not in profiles["profiles"]:
        click.echo(f"Profile '{profile}' not found. Run 'setup' first.")
        return

    kis_dir = Path(profiles["profiles"][profile]["kis_dir"])
    client_dir = Path(profiles["profiles"][profile]["client_dir"])
    excluded_files = set(profiles["excluded_files"])
    excluded_dirs = set(profiles["excluded_dirs"])

    click.echo(f"Comparing files between:\n  KIS: {kis_dir}\n  {profile}: {client_dir}\n")

    def should_ignore(path, base_dir):
        """Check if a file or directory should be ignored."""
        try:
            relative_path = path.relative_to(base_dir)
        except ValueError:
            relative_path = path.name

        if path.name in excluded_files or any(path.name.startswith(prefix) for prefix in excluded_files):
            click.echo(click.style(f"Skipping ignored file: {relative_path}", fg="yellow"))
            return True

        if any(excluded_dir in path.parts for excluded_dir in excluded_dirs):
            click.echo(click.style(f"Skipping ignored directory: {relative_path}", fg="yellow"))
            return True

        return False

    # Sync files from KIS to Client
    for kis_file in kis_dir.rglob("*"):
        if should_ignore(kis_file, kis_dir):
            continue

        if kis_file.is_file():
            relative_path = kis_file.relative_to(kis_dir)
            client_file = client_dir / relative_path

            if not client_file.exists():
                if click.confirm(f"File missing in {profile} SharePoint: {relative_path}. Copy from KIS SharePoint?"):
                    os.makedirs(client_file.parent, exist_ok=True)
                    shutil.copy2(kis_file, client_file)
                    click.echo(f"Copied {relative_path} to {profile} SharePoint.")
            else:
                kis_mtime = kis_file.stat().st_mtime
                client_mtime = client_file.stat().st_mtime

                if kis_mtime == client_mtime:
                    click.echo(click.style(f"File {relative_path} is the same version in both SharePoints. Skipping.", fg="cyan"))
                    continue

                # Determine which file is newer
                if kis_mtime > client_mtime:
                    source_file, target_file, source_name, target_name = kis_file, client_file, "KIS", profile
                else:
                    source_file, target_file, source_name, target_name = client_file, kis_file, profile, "KIS"

                if not filecmp.cmp(kis_file, client_file, shallow=False):
                    click.echo("Differences detected")
                    show_file_diff(kis_file, client_file, source_name)

                if click.confirm(f"Copy newer version from {click.style(source_name, fg='green')} SharePoint to {click.style(target_name, fg='yellow')} SharePoint?"):
                    shutil.copy2(source_file, target_file)
                    click.echo(f"Updated {relative_path} with newer version from {source_name} SharePoint.")

    # Sync files from Client to KIS
    for client_file in client_dir.rglob("*"):
        if should_ignore(client_file, client_dir):
            continue

        if client_file.is_file():
            relative_path = client_file.relative_to(client_dir)
            kis_file = kis_dir / relative_path

            if not kis_file.exists():
                if click.confirm(f"File missing in KIS SharePoint: {relative_path}. Copy from {profile} SharePoint?"):
                    os.makedirs(kis_file.parent, exist_ok=True)
                    shutil.copy2(client_file, kis_file)
                    click.echo(f"Copied {relative_path} to KIS SharePoint.")

@click.command()
@click.argument("dir_name")
def exclude_dir(dir_name):
    "Add a directory name to the exclusion list"
    profiles = load_profiles()
    if dir_name not in profiles["excluded_dirs"]:
        profiles["excluded_dirs"].append(dir_name)
        save_profiles(profiles)
        click.echo(f"Directory '{dir_name}' added to exclusion list.")
    else:
        click.echo(f"Directory '{dir_name}' is already excluded.")

@click.command()
@click.argument("file_name")
def exclude_file(file_name):
    "Add a file name to the exclusion list"
    profiles = load_profiles()
    if file_name not in profiles["excluded_files"]:
        profiles["excluded_files"].append(file_name)
        save_profiles(profiles)
        click.echo(f"File '{file_name}' added to exclusion list.")
    else:
        click.echo(f"File '{file_name}' is already excluded.")

cli.add_command(setup)
cli.add_command(sync)
cli.add_command(exclude_dir)
cli.add_command(exclude_file)

if __name__ == "__main__":
    cli()
