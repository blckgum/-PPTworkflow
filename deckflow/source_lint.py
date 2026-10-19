"""源码卫生检查，不依赖任何第三方包或其他引擎模块。"""
from pathlib import Path
import io
import re
import tokenize


def lint_sources(root: Path) -> list[str]:
    errors = []
    for path in sorted(root.glob("*.py")):
        try:
            source = path.read_text(encoding="utf-8")
            compile(source, str(path), "exec")
            # 注释和文档允许中文；运行时字符串必须使用转义或内容文件。
            tokens = list(tokenize.generate_tokens(io.StringIO(source).readline))
            docstring_lines = set()
            import ast
            tree = ast.parse(source)
            for node in ast.walk(tree):
                body = getattr(node, "body", None)
                if isinstance(body, list) and body and isinstance(body[0], ast.Expr):
                    value = body[0].value
                    if isinstance(value, ast.Constant) and isinstance(value.value, str):
                        docstring_lines.add(value.lineno)
            for token in tokens:
                string_types = {tokenize.STRING, getattr(tokenize, "FSTRING_MIDDLE", -1)}
                if token.type not in string_types or token.start[0] in docstring_lines:
                    continue
                if re.search(r"[\u3400-\u9fff]", token.string):
                    errors.append(f"{path}:{token.start[0]} "
                                  "\u4e2d\u6587\u8fd0\u884c\u65f6\u5b57\u7b26\u4e32\u5fc5\u987b\u79fb\u81f3\u5185\u5bb9\u6587\u4ef6\u6216\u4f7f\u7528 Unicode \u8f6c\u4e49")
        except (SyntaxError, tokenize.TokenError, UnicodeError) as exc:
            errors.append(f"{path}:{getattr(exc, 'lineno', 1)} {exc}")
    return errors
