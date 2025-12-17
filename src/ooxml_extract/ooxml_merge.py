import click
import subprocess
from pathlib import Path
import tempfile
from .ooxml_package import extract_ooxml, pack_ooxml


def run(args, cwd):
    try:
        subprocess.check_call(
            args,
            cwd=cwd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except subprocess.CalledProcessError as e:
        click.echo(f"Error on comand: {' '.join(args)}", err=True)
        raise e
    

def ensure_git_available():
    try:
        subprocess.run(["git", "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    except Exception as e:
        raise click.ClickException("Git ist nicht verfügbar (git --version fehlgeschlagen).") from e


def automerge(ooxml_original: Path, ooxml_a: Path, ooxml_b: Path, ooxml_merged: Path, force: bool):
    ensure_git_available()

    if ooxml_merged.exists() and not force:
        raise click.ClickException(f"File exists: '{ooxml_merged}'. Use --force, to overwrite.")

    with tempfile.TemporaryDirectory(prefix="folder-merge-git-") as repo:
        #repo.mkdir(parents=True, exist_ok=True)
        repo = Path(repo)

        # 1) Git-Repo initialisieren
        run(["git", "init", "-q"], cwd=repo)
        run(["git", "config", "user.email", "mail@local"], cwd=repo)
        run(["git", "config", "user.name", "Name"], cwd=repo)

        # 2) Original entpacken und als ersten Commit hinzufügen
        extract_ooxml(ooxml_original, repo / "ooxml", overwrite=True, prettify=True)
        run(["git", "add", "."], cwd=repo)
        run(["git", "commit", "-m", "Original"], cwd=repo)

        # 3) Version A entpacken und als Branch hinzufügen
        run(["git", "checkout", "-b", "branch_a"], cwd=repo)
        extract_ooxml(ooxml_a, repo / "ooxml", overwrite=True, prettify=True)
        run(["git", "add", "."], cwd=repo)
        run(["git", "commit", "-m", "Version A"], cwd=repo)

        # 4) Version B entpacken und als Branch hinzufügen
        run(["git", "checkout", "master"], cwd=repo)
        run(["git", "checkout", "-b", "branch_b"], cwd=repo)
        extract_ooxml(ooxml_b, repo / "ooxml", overwrite=True, prettify=True)
        run(["git", "add", "."], cwd=repo)
        run(["git", "commit", "-m", "Version B"], cwd=repo)

        # 5) Merge
        run(["git", "checkout", "master"], cwd=repo)
        run(["git", "merge", "branch_a", "-X", "theirs", "--no-edit"], cwd=repo)
        run(["git", "merge", "branch_b", "-X", "theirs", "--no-edit"], cwd=repo)

        # 6) Packen
        pack_ooxml(repo / "ooxml", ooxml_merged, overwrite=True)
        