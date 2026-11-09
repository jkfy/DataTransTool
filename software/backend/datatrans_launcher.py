"""PyInstaller 入口：使用绝对导入，避免 __main__.py 的相对导入在打包后失效。"""

from datatrans.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
