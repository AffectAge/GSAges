"""Prepare a downloaded Apps Script project without discarding unknown code."""
import json
import pathlib
import shutil
import subprocess
import sys
import tempfile


def prepare(source, target):
    scripts = sorted(source.glob("*.gs"))
    if not scripts:
        raise ValueError("No repository .gs files found")
    names = {path.stem for path in scripts}
    # The linked project uses this historical name for the same engine.
    aliases = {"GAME_TURN_ENGINE": "GameTurnEngine"}
    legacy_names = {old for old, new in aliases.items() if new in names}
    config = json.loads((target / ".clasp.json").read_text(encoding="utf-8"))
    expected = "1ArxSXXuUjy5th8-M2DQ7Gl9TOD6QFXmxk1XZMkHyaG-B-K7Mhq8qyfrS"
    if config.get("scriptId") != expected:
        raise ValueError("Wrong Apps Script project; no files changed")
    manifest = target / "appsscript.json"
    if not manifest.is_file():
        raise ValueError("Downloaded project has no manifest")
    json.loads(manifest.read_text(encoding="utf-8"))
    unknown = [path.relative_to(target).as_posix()
               for path in target.rglob("*")
               if path.suffix in {".js", ".gs"} and path.is_file()
               and (path.parent != target or path.stem not in names | legacy_names)]
    if unknown:
        raise ValueError("Project contains code absent from GitHub. Import it into the repository "
                         "or reconcile it before syncing: " + ", ".join(sorted(unknown)))
    with tempfile.TemporaryDirectory() as tmp:
        for path in scripts:
            check = pathlib.Path(tmp) / (path.stem + ".js")
            shutil.copyfile(path, check)
            subprocess.run(["node", "--check", str(check)], check=True)
        # Apps Script files share a global scope: check combined declarations too.
        combined = pathlib.Path(tmp) / "combined.js"
        combined.write_text("\n".join(path.read_text(encoding="utf-8") for path in scripts),
                            encoding="utf-8")
        subprocess.run(["node", "--check", str(combined)], check=True)
    # clasp clone may produce .js rather than .gs. Avoid uploading both.
    for name in legacy_names:
        for suffix in (".js", ".gs"):
            existing = target / (name + suffix)
            if existing.exists():
                existing.unlink()
    for path in scripts:
        for suffix in (".js", ".gs"):
            existing = target / (path.stem + suffix)
            if existing.exists():
                existing.unlink()
        shutil.copyfile(path, target / path.name)
    print(f"Prepared {len(scripts)} scripts; remote manifest and HTML preserved.")


if __name__ == "__main__":
    try:
        prepare(pathlib.Path(sys.argv[1]).resolve(), pathlib.Path(sys.argv[2]).resolve())
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        print(f"Sync stopped: {error}", file=sys.stderr)
        sys.exit(1)
