"""中文命令行；任何失败返回非零退出状态。"""
import argparse
from pathlib import Path
import shutil
import sys

from .content import WorkflowError, load_course
from .strings import tr
from .theme import get_theme


def course_path(root, name):
    path = root / "courses" / name
    if (name in {"", ".", ".."} or any(char in name for char in "/\\") or
            not path.resolve().is_relative_to((root / "courses").resolve())):
        raise WorkflowError("INPUT", tr("bad_course_name"))
    return path


def main(root: Path):
    parser = argparse.ArgumentParser(description=tr("cli_description"))
    parser.add_argument("course", nargs="?", help=tr("cli_course"))
    parser.add_argument("--all", action="store_true", help=tr("cli_all"))
    parser.add_argument("--check", action="store_true", help=tr("cli_check"))
    parser.add_argument("--new", metavar="NAME", help=tr("cli_new"))
    args = parser.parse_args()
    if sum(bool(item) for item in (args.course, args.all, args.new)) != 1 or (args.new and args.check):
        parser.error(tr("cli_args"))
    try:
        if args.new:
            path = course_path(root, args.new)
            if path.exists():
                raise WorkflowError("INPUT", tr("course_exists", path=path))
            shutil.copytree(root / "courses" / "_template", path)
            print(tr("new_done", path=path))
            return 0
        names = [args.course] if args.course else sorted(
            path.name for path in (root / "courses").iterdir()
            if path.is_dir() and not path.name.startswith("_") and (path / "course.txt").is_file())
        if not names:
            raise WorkflowError("INPUT", tr("no_courses"))
        from .build_docx import build_docx
        from .build_pptx import build_pptx
        from .checks import check_outputs, passed, preflight, report
        success = True
        for name in names:
            try:
                path = course_path(root, name)
                if not path.is_dir():
                    raise WorkflowError("INPUT", tr("course_missing", path=path))
                print(tr("check_start" if args.check else "build_start", name=name))
                course = load_course(path)
                if not args.check:
                    checks = preflight(course, root / "deckflow")
                    if not passed(checks):
                        report(checks)
                        success = False
                        continue
                    theme = get_theme(course.meta["theme"])
                    build_pptx(course, theme)
                    build_docx(course, theme)
                checks = check_outputs(course, root / "deckflow")
                report(checks)
                if passed(checks):
                    if not args.check:
                        print(tr("build_done", path=course.pptx_path))
                        print(tr("build_done", path=course.docx_path))
                else:
                    success = False
            except WorkflowError as exc:
                print(f"[{tr('fail')}] {exc}")
                success = False
        return 0 if success else 1
    except WorkflowError as exc:
        print(f"[{tr('fail')}] {exc}")
        return 1
    except ImportError as exc:
        print(tr("dependency_missing", error=exc), file=sys.stderr)
        return 1
    except Exception as exc:
        print(tr("unexpected", error=exc), file=sys.stderr)
        return 1
