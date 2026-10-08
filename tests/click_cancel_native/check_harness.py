"""Focused controlled harness checks. No Access/COM instances or fault injection."""
import sys
sys.dont_write_bytecode = True
import copy
import json
from pathlib import Path
from contextlib import contextmanager
import uuid
import unittest
from unittest.mock import Mock, patch
import native_cancel as driver
import run_click_cancel as runner
from msaccess_vcs_mcp.dialog_recovery import CLICK_UNCERTAIN, Win32Backend

HERE = runner.ARTIFACTS
HERE.mkdir(parents=True, exist_ok=True)


@contextmanager
def controlled_files():
    # Retain controlled receipts; no recursive deletion or native Access work.
    path = HERE / ('checks-' + uuid.uuid4().hex)
    path.mkdir()
    yield path


def dialog():
    return dict(hwnd=10, pid=22, visible=True, owner=9,
        texts=[driver.PROMPT], buttons=[dict(hwnd=11, text=s, path=(n,))
            for n, s in enumerate(('Yes', 'No', 'Cancel'), 1)], **{'class': 'NUIDialog'})


class HarnessChecks(unittest.TestCase):
    def test_target_match_rejects_wrong_process_prompt_and_button(self):
        self.assertEqual(driver.validate_dialog(dialog(), 22)['text'], 'Cancel')
        for key, value in (('pid', 23), ('texts', ['Other form']), ('buttons', []),
                           ('class', 'OtherWindow'), ('visible', False)):
            with self.subTest(key=key):
                d = dialog()
                d[key] = value
                with self.assertRaises(AssertionError):
                    driver.validate_dialog(d, 22)

    def test_uncertain_click_recorded_and_never_replayed(self):
        with controlled_files() as directory:
            work = Path(directory)
            pending = dict(case='export', pid=22, creation_FILETIME='134359635370106083')
            with patch.object(driver, 'observe', return_value=dialog()), \
                 patch('win32process.GetWindowThreadProcessId', return_value=(1, 22)), \
                 patch('win32gui.IsChild', return_value=True), \
                 patch.object(Win32Backend, 'click', return_value=CLICK_UNCERTAIN) as dispatch:
                with self.assertRaises(AssertionError):
                    driver.click(work, pending, lambda: None)
                receipt = json.loads((work / 'export-click.json').read_text())
                self.assertEqual(receipt['status'], 'uncertain')
                self.assertEqual(receipt['pending']['creation_FILETIME'], pending['creation_FILETIME'])
                with self.assertRaises(AssertionError):
                    driver.click(work, pending, lambda: None)
                self.assertEqual(dispatch.call_count, 1)

    def test_changed_dialog_causes_no_dispatch(self):
        with controlled_files() as directory:
            changed = dialog()
            changed['hwnd'] = 100
            with patch.object(driver, 'observe', side_effect=[dialog(), changed]), \
                 patch.object(Win32Backend, 'click') as dispatch:
                with self.assertRaises(AssertionError):
                    driver.click(Path(directory), dict(case='import', pid=22), lambda: None)
                dispatch.assert_not_called()

    def test_uncertain_ownership_causes_no_observation_or_click(self):
        with patch.object(driver, 'observe') as observe:
            with self.assertRaises(RuntimeError):
                driver.click(HERE, dict(pid=22), Mock(side_effect=RuntimeError('identity uncertain')))
            observe.assert_not_called()

    def test_dialog_wait_is_bounded(self):
        with patch.object(driver, 'observe', return_value=None):
            with self.assertRaises(AssertionError):
                driver.click(HERE, dict(pid=22), lambda: None, timeout=0)

    def test_readiness_and_completion_wait_are_bounded(self):
        with self.assertRaises(TimeoutError):
            runner.wait_for(lambda: False, 0, Mock())
        with self.assertRaises(RuntimeError):
            runner.wait_for(lambda: False, 1, Mock(poll=Mock(return_value=1)))

    def test_final_pass_requires_preservation_cleanup_and_behavior(self):
        good = dict(outcome='behavior_passed', installed_unchanged=True,
                    ownership=dict(original_handle_exit_confirmed=True))
        self.assertTrue(runner.final_pass(good, True, True, []))
        self.assertFalse(runner.final_pass(good, False, True, []))
        self.assertFalse(runner.final_pass(good, True, False, []))
        self.assertFalse(runner.final_pass(good, True, True, ['cleanup failure']))
        failed = copy.deepcopy(good)
        failed['ownership']['original_handle_exit_confirmed'] = False
        self.assertFalse(runner.final_pass(failed, True, True, []))
        failed['outcome'] = 'failed'
        self.assertFalse(runner.final_pass(failed, True, True, []))

    def test_environment_is_restored_on_exception(self):
        import os
        before = dict(os.environ)
        with patch.object(runner, 'run', side_effect=RuntimeError('controlled failure')):
            with self.assertRaises(RuntimeError):
                runner.entry()
        self.assertEqual(dict(os.environ), before)

    def test_cleanup_failure_writes_failed_receipt_and_nonzero_exit(self):
        with controlled_files() as work:
            receipt = dict(errors=['controlled cleanup failure'], installed_unchanged=True,
                           original_handle_exit_confirmed=True, evidence_kind='controlled; no native fault')
            good = dict(outcome='behavior_passed', installed_unchanged=True,
                        ownership=dict(original_handle_exit_confirmed=True))
            self.assertEqual(runner.finish(work, receipt, good), 1)
            self.assertEqual(json.loads((work / 'result.json').read_text())['outcome'], 'failed')


if __name__ == '__main__':
    unittest.main(verbosity=2)
