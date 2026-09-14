# 8D Issue Automation v3.0.9 DEBUG
# Diagnostic launcher only. It does not change the v3.0.8 processing logic.
# It records the exact Python exception traceback to error_v309.log even when
# the GUI catches the exception and displays only the exception message.
import sys
import traceback
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
LOG = APP_DIR.parent / 'error_v309.log'


def trace_exceptions(frame, event, arg):
    if event == 'exception':
        try:
            exc_type, exc_value, exc_tb = arg
            with LOG.open('a', encoding='utf-8') as f:
                f.write('\n' + '=' * 90 + '\n')
                f.write('EXCEPTION EVENT\n')
                f.write(f'Type: {exc_type.__name__}\n')
                f.write(f'Message: {exc_value!r}\n')
                f.write('Traceback:\n')
                f.writelines(traceback.format_tb(exc_tb))
                f.write('=' * 90 + '\n')
        except Exception:
            pass
    return trace_exceptions

try:
    LOG.write_text('', encoding='utf-8')
except Exception:
    pass

# Install tracing before loading the application so import-time exceptions are also captured.
sys.settrace(trace_exceptions)

try:
    import main_v308 as app
    app.base.App().mainloop()
except Exception as exc:
    with LOG.open('a', encoding='utf-8') as f:
        f.write('\n' + '=' * 90 + '\nUNHANDLED TOP-LEVEL EXCEPTION\n')
        traceback.print_exc(file=f)
        f.write('=' * 90 + '\n')
    raise
