from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
try:
    from desktop_query import main
    main()
except Exception as error:
    message = str(error) if isinstance(error, RuntimeError) and str(error).startswith('查询库暂无资料') else f'无法启动规则与例句查询：\n{error}\n\n请检查Python、pywebview及本系统所需的窗口组件。'
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, '规则与例句查询', 16)
    else:
        print(message, file=sys.stderr)
    sys.exit(1)
