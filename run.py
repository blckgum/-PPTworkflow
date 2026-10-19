"""课程构建入口；先检查引擎源码，再导入业务模块。"""
from pathlib import Path
import sys


def main():
    root = Path(__file__).resolve().parent
    # 语法损坏的模块也必须以 V7 清晰报错，不能先导入它。
    for path in sorted((root / "deckflow").glob("*.py")):
        try:
            compile(path.read_text(encoding="utf-8"), str(path), "exec")
        except (SyntaxError, UnicodeError) as exc:
            print("[\u5931\u8d25] V7 {}:{} {}".format(
                path, getattr(exc, "lineno", 1), str(exc)))
            return 1
    from deckflow.source_lint import lint_sources
    failures = lint_sources(root / "deckflow")
    if failures:
        for failure in failures:
            print("[\u5931\u8d25] V7 " + failure)
        return 1
    from deckflow.cli import main as cli_main
    return cli_main(root)


if __name__ == "__main__":
    sys.exit(main())
