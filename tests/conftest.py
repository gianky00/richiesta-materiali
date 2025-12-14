import sys
import os
import types
from unittest.mock import MagicMock
import pytest

# Add src to path to ensure modules can be found
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# --- Mock Windows-specific modules GLOBALLY before any test imports code ---

module_names = [
    'pythoncom',
    'win32com',
    'win32com.client',
    'win32api',
    'win32gui',
    'win32con'
]

for name in module_names:
    if name not in sys.modules:
        m = types.ModuleType(name)
        # Add basic attributes often accessed
        if name == 'win32com.client':
            m.Dispatch = MagicMock()
            m.DispatchEx = MagicMock()
        elif name == 'pythoncom':
            m.CoInitialize = MagicMock()
            m.CoUninitialize = MagicMock()

        sys.modules[name] = m

# Mock tkinter completely for headless testing
# We force mock it even if it exists, to avoid _tkinter.TclError: no display
tkinter = MagicMock()
sys.modules['tkinter'] = tkinter
sys.modules['tkinter.ttk'] = MagicMock()
sys.modules['tkinter.messagebox'] = MagicMock()
sys.modules['tkinter.filedialog'] = MagicMock()
# Also catch _tkinter if it tries to load low level
sys.modules['_tkinter'] = MagicMock()


@pytest.fixture(autouse=True)
def mock_settings(monkeypatch):
    """
    Setup common mocks for all tests.
    """
    # Prevent webbrowser.open from actually opening browsers
    monkeypatch.setattr("webbrowser.open", lambda url: print(f"Mock open browser: {url}"))

    # Mock messagebox to avoid hanging
    pass
