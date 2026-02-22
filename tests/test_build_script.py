from pathlib import Path
import unittest


class BuildScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.script = Path('build_exe.bat').read_text(encoding='utf-8')

    def test_has_log_file_output(self):
        self.assertIn('set "LOG_FILE=%OUTPUT_DIR%\\build.log"', self.script)
        self.assertIn('Build failed. See log for details', self.script)

    def test_has_pause_control(self):
        self.assertIn('set "KEEP_WINDOW_OPEN=1"', self.script)
        self.assertIn('if /I "%~1"=="--no-pause" set "KEEP_WINDOW_OPEN=0"', self.script)
        self.assertIn('if "%KEEP_WINDOW_OPEN%"=="1" (', self.script)

    def test_keeps_win32com_packaging_flags(self):
        self.assertIn('python -m PyInstaller --noconfirm --clean --onefile --windowed', self.script)
        self.assertIn('--hidden-import pythoncom', self.script)
        self.assertIn('--hidden-import pywintypes', self.script)
        self.assertIn('--collect-submodules win32com', self.script)


if __name__ == '__main__':
    unittest.main()
