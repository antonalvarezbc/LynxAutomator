from types import SimpleNamespace
import unittest
from unittest.mock import patch
import lynx_dialogs


class DialogTests(unittest.TestCase):
    def test_linux_filters_and_remembered_directory(self):
        with patch.object(lynx_dialogs.sys, 'platform', 'linux'), \
                patch.object(lynx_dialogs.shutil, 'which', return_value='/usr/bin/zenity'), \
                patch.object(lynx_dialogs, '_last_directory', '/tmp'), \
                patch.object(lynx_dialogs.subprocess, 'run', return_value=SimpleNamespace(returncode=0, stdout='/tmp/data/PHOTO.XLSX\n')) as run:
            result = lynx_dialogs.askopenfilename(filetypes=[('Excel', '*.xlsx;*.xlsm')])
            self.assertEqual(result, '/tmp/data/PHOTO.XLSX')
            self.assertEqual(lynx_dialogs._last_directory, '/tmp/data')
            self.assertIn('--file-filter=Excel | *.xlsx *.XLSX *.xlsm *.XLSM', run.call_args.args[0])
            self.assertIn('--file-filter=All files | *', run.call_args.args[0])

    def test_native_cancel_does_not_open_second_dialog(self):
        with patch.object(lynx_dialogs.sys, 'platform', 'linux'), \
                patch.object(lynx_dialogs.shutil, 'which', return_value='/usr/bin/zenity'), \
                patch.object(lynx_dialogs.subprocess, 'run', return_value=SimpleNamespace(returncode=1, stdout='')):
            self.assertEqual(lynx_dialogs.askopenfilename(), '')

    def test_native_failure_falls_back_to_tk(self):
        from unittest.mock import Mock
        fallback = Mock(return_value='/tmp/example.xlsx')
        tkinter = SimpleNamespace(filedialog=SimpleNamespace(askopenfilename=fallback))
        with patch.dict('sys.modules', {'tkinter': tkinter}), \
                patch.object(lynx_dialogs.sys, 'platform', 'linux'), \
                patch.object(lynx_dialogs, '_last_directory', '/tmp'), \
                patch.object(lynx_dialogs.shutil, 'which', return_value='/usr/bin/zenity'), \
                patch.object(lynx_dialogs.subprocess, 'run', side_effect=OSError):
            self.assertEqual(lynx_dialogs.askopenfilename(), '/tmp/example.xlsx')
            fallback.assert_called_once_with(initialdir='/tmp')
