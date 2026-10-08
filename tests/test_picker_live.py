"""Native no-argument BuildAs pickers, using disposable STA workers."""
import sys
import pytest

from tests.picker_native import run_case, sentinel as native_sentinel

pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(sys.platform != 'win32', reason='Requires Microsoft Access and Windows dialogs'),
]


@pytest.fixture(scope='module')
def sentinel(tmp_path_factory):
    with native_sentinel(tmp_path_factory.mktemp('unrelated')) as identity:
        yield identity


def test_build_as_actual_pickers_create_selected_database(tmp_path, sentinel):
    result = run_case(tmp_path, 'success')
    assert result['passed'] and result['unrelated']


@pytest.mark.parametrize('case', ['cancel_source', 'cancel_output'])
def test_build_as_cancel_actual_picker_preserves_host(tmp_path, case, sentinel):
    result = run_case(tmp_path, case)
    assert result['passed'] and result['unrelated']


def test_build_as_invalid_source_shows_error_and_preserves_host(tmp_path, sentinel):
    result = run_case(tmp_path, 'invalid_source')
    assert result['passed'] and result['unrelated']
