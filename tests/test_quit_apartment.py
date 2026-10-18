"""Recovery must acquire its own apartment proxy and stay bounded on failure."""
import threading
import time
import gc
import weakref

import pytest

from msaccess_vcs_mcp.access_com import connection


@pytest.fixture
def apartments(monkeypatch):
    events=[]
    release=threading.Event()
    state=threading.local()

    class App:
        def __init__(self, failure=None):
            self.owner=threading.get_ident()
            self.failure=failure
            self._oleobj_=self
        def CloseCurrentDatabase(self):
            assert self.owner==threading.get_ident(), 'wrong apartment'
            events.append(('close',self.owner))
            if self.failure=='close': raise RuntimeError('close failed')
        def Quit(self):
            assert self.owner==threading.get_ident(), 'wrong apartment'
            events.append(('quit',self.owner))
            if self.failure=='quit': raise RuntimeError('quit failed')
            if self.failure=='blocked': release.wait(2)

    def initialize():
        state.initialized=True
        events.append(('initialize',threading.get_ident()))
    def uninitialize():
        assert state.initialized
        state.initialized=False
        events.append(('uninitialize',threading.get_ident()))
    def marshal(iid, app):
        assert app.owner==threading.get_ident()
        events.append(('marshal',threading.get_ident()))
        return {'failure':app.failure,'creator':app.owner}
    def unmarshal(stream,iid):
        assert state.initialized
        assert stream['creator']!=threading.get_ident()
        events.append(('unmarshal',threading.get_ident()))
        return App(stream['failure'])

    monkeypatch.setattr(connection,'COM_AVAILABLE',True)
    monkeypatch.setattr(connection.pythoncom,'CoInitialize',initialize)
    monkeypatch.setattr(connection.pythoncom,'CoUninitialize',uninitialize)
    monkeypatch.setattr(connection.pythoncom,'CoMarshalInterThreadInterfaceInStream',marshal)
    monkeypatch.setattr(connection.pythoncom,'CoGetInterfaceAndReleaseStream',unmarshal)
    monkeypatch.setattr(connection.win32com.client,'Dispatch',lambda value:value)
    yield App,events,release
    release.set()
    limit=time.monotonic()+1
    while any(t.name=='vcs-access-quit' and t.is_alive() for t in threading.enumerate()) and time.monotonic()<limit:
        time.sleep(.005)


@pytest.mark.parametrize('failure,expected',[(None,True),('close',True),('quit',False)])
def test_quit_uses_worker_proxy_and_balances_apartment(apartments,failure,expected):
    App,events,_=apartments
    creator=threading.get_ident()
    assert connection._quit_with_timeout(App(failure),.5) is expected
    names=[event[0] for event in events]
    assert names[:3]==['marshal','initialize','unmarshal']
    assert names[-1]=='uninitialize'
    assert all(tid!=creator for name,tid in events if name in ['close','quit'])
    assert 'quit' in names


def test_blocked_call_does_not_block_independent_inspection(apartments):
    App,events,release=apartments
    start=time.monotonic()
    assert connection._quit_with_timeout(App('blocked'),.02) is False
    assert time.monotonic()-start<.5
    inspected=threading.Event()
    inspector=threading.Thread(target=inspected.set)
    inspector.start(); inspector.join(.2)
    assert inspected.is_set()
    assert any(name=='unmarshal' for name,_ in events)
    release.set()


def test_marshal_failure_is_bounded_and_does_not_use_caller_proxy(apartments,monkeypatch):
    App,events,_=apartments
    def fail(*args): raise RuntimeError('marshal failed')
    monkeypatch.setattr(connection.pythoncom,'CoMarshalInterThreadInterfaceInStream',fail)
    assert connection._quit_with_timeout(App(),.02) is False
    assert not events


def test_timeout_does_not_retain_caller_proxy(apartments):
    App,_,release=apartments
    app=App('blocked')
    reference=weakref.ref(app)
    assert connection._quit_with_timeout(app,.02) is False
    del app
    gc.collect()
    assert reference() is None
    release.set()


def test_failed_initialization_releases_unused_marshal_on_owner(apartments,monkeypatch):
    App,events,_=apartments
    creator=threading.get_ident()
    def fail(): raise RuntimeError('init failed')
    monkeypatch.setattr(connection.pythoncom,'CoInitialize',fail)
    monkeypatch.setattr(connection.pythoncom,'CoReleaseMarshalData',lambda stream:events.append(('release',threading.get_ident())))
    assert connection._quit_with_timeout(App(),.5) is False
    assert events==[('marshal',creator),('release',creator)]


def test_failed_thread_start_releases_unused_marshal(apartments,monkeypatch):
    App,events,_=apartments
    creator=threading.get_ident()
    class FailedThread:
        def __init__(self,**kwargs): pass
        def start(self): raise RuntimeError('start failed')
    monkeypatch.setattr(connection.threading,'Thread',FailedThread)
    monkeypatch.setattr(connection.pythoncom,'CoReleaseMarshalData',lambda stream:events.append(('release',threading.get_ident())))
    assert connection._quit_with_timeout(App(),.5) is False
    assert events==[('marshal',creator),('release',creator)]
