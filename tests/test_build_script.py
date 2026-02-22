from pathlib import Path
import re
import unittest


class BuildScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        raw = Path('build_exe.bat').read_text(encoding='utf-8')
        cls.script = raw.replace('\r\n', '\n')

    def test_has_log_file_output(self):
        self.assertIn('set "LOG_FILE=%OUTPUT_DIR%\\build.log"', self.script)
        self.assertIn('Build failed. See log for details', self.script)
        self.assertIn('Get-Content', self.script)

    def test_has_pause_control(self):
        self.assertRegex(self.script, r'set\s+"KEEP_WINDOW_OPEN=1"')
        self.assertRegex(self.script, r'if\s+/I\s+"%~1"=="--no-pause"\s+set\s+"KEEP_WINDOW_OPEN=0"')
        self.assertRegex(self.script, r'if\s+"%KEEP_WINDOW_OPEN%"=="1"\s*\(')

    def test_keeps_win32com_packaging_flags(self):
        self.assertIn('--hidden-import pythoncom', self.script)
        self.assertIn('--hidden-import pywintypes', self.script)
        self.assertIn('--collect-submodules win32com', self.script)

    def test_no_premature_goto_eof_before_dist_fallback(self):
        output_loop = self.script.find('dir /b /s "%OUTPUT_DIR%\\WordToPdfConverter*.exe"')
        dist_loop = self.script.find('dir /b /s "%SCRIPT_DIR%dist\\WordToPdfConverter*.exe"')
        self.assertNotEqual(output_loop, -1)
        self.assertNotEqual(dist_loop, -1)
        self.assertLess(output_loop, dist_loop)
        between = self.script[output_loop:dist_loop]
        self.assertNotIn('goto :eof', between)


if __name__ == '__main__':
    unittest.main()
