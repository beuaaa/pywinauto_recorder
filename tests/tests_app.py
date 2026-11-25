# -*- coding: utf-8 -*-
"""
Tests of the Pywinauto Recorder application (pywinauto_recorder/app.py).

They need no application under test and no user interaction, so they run in the CI.
"""

import pathlib

import pytest

import pywinauto_recorder
from pywinauto_recorder import app, recorder

PACKAGE_DIR = pathlib.Path(pywinauto_recorder.__file__).resolve().parent


def test_app_has_main():
	"""Tests that the application is started by pywinauto_recorder.app.main."""
	assert callable(app.main)


def test_icons_are_in_the_package():
	"""Tests that the icons are in the package, where the application and the recorder search them."""
	icons_dir = PACKAGE_DIR / "Icons"
	for name in ("IconPyRec.ico", "record.ico", "stop.ico", "play.ico", "folder.ico", "search.ico", "light-on.ico",
	             "settings.ico", "check.ico", "cross.ico", "power.ico", "help.ico", "comments.ico", "favourite.ico",
	             "paste.ico"):
		assert (icons_dir / name).is_file(), name
	assert pathlib.Path(recorder.IconSet.path_icons).resolve() == PACKAGE_DIR


def test_command_pywinauto_recorder():
	"""Tests that installing the package installs the command 'pywinauto-recorder'."""
	from importlib import metadata
	try:
		distribution = metadata.distribution("pywinauto_recorder")
	except metadata.PackageNotFoundError:
		pytest.skip("The package is not installed")
	entry_points = [entry_point for entry_point in distribution.entry_points
	                if entry_point.group == "console_scripts" and entry_point.name == "pywinauto-recorder"]
	assert len(entry_points) == 1
	assert entry_points[0].value == "pywinauto_recorder.app:main"


def test_python_m_pywinauto_recorder():
	"""Tests that the package can be started with 'python -m pywinauto_recorder'."""
	assert (PACKAGE_DIR / "__main__.py").is_file()
	assert "from pywinauto_recorder.app import main" in (PACKAGE_DIR / "__main__.py").read_text()


################################################################################################
#                                     Modes of the recorder                                    #
################################################################################################

@pytest.mark.parametrize("mode, value", [
	(recorder.RecorderMode.INITIALIZING, "Initializing"),
	(recorder.RecorderMode.RECORD, "Record"),
	(recorder.RecorderMode.PLAY, "Play"),
	(recorder.RecorderMode.INFO, "Info"),
	(recorder.RecorderMode.STOP, "Stop"),
	(recorder.RecorderMode.QUIT, "Quit"),
])
def test_recorder_mode_is_equal_to_its_value(mode, value):
	"""Tests that a mode is equal to its former string value, so that code comparing it with a string keeps working."""
	assert isinstance(mode, str)
	assert mode == value and value == mode
	assert mode in (value,)
	assert recorder.RecorderMode(value) is mode


class _FakeRecorder:
	"""Has the attribute used by the 'mode' property of the Recorder, without starting a recorder."""
	_mode = None


@pytest.mark.parametrize("value", ["Record", recorder.RecorderMode.RECORD])
def test_recorder_mode_setter_accepts_a_string_or_a_mode(value):
	"""Tests that the mode of the recorder can be set with a RecorderMode or with its value."""
	fake_recorder = _FakeRecorder()
	recorder.Recorder.mode.fset(fake_recorder, value)
	assert recorder.Recorder.mode.fget(fake_recorder) is recorder.RecorderMode.RECORD


def test_recorder_mode_setter_refuses_an_unknown_mode():
	"""Tests that the mode of the recorder cannot be set with an unknown value."""
	with pytest.raises(ValueError):
		recorder.Recorder.mode.fset(_FakeRecorder(), "Pause")


def test_recorder_mode_is_exported_by_the_package():
	"""Tests that RecorderMode can be imported from the package."""
	assert pywinauto_recorder.RecorderMode is recorder.RecorderMode


################################################################################################
#                                 Command line and file replay                                 #
################################################################################################

import ctypes
import subprocess
import sys
import types


def test_parse_arguments_without_argument():
	"""Tests the command line without argument: the recorder starts with its splash screen."""
	args = app.parse_arguments([])
	assert args.filename == ""
	assert args.no_splash_screen is False


def test_parse_arguments_with_a_script_and_no_splash_screen():
	"""Tests the command line with a script to replay and --no_splash_screen."""
	args = app.parse_arguments(["C:\\Scripts\\recorded file.py", "--no_splash_screen"])
	assert args.filename == "C:\\Scripts\\recorded file.py"
	assert args.no_splash_screen is True


@pytest.mark.parametrize("argv0, program_name", [
	("C:\\Python\\Lib\\site-packages\\pywinauto_recorder\\__main__.py", "python -m pywinauto_recorder"),
	("C:\\Program Files\\Pywinauto recorder\\pywinauto_recorder.exe", "pywinauto_recorder.exe"),
	("pywinauto-recorder", "pywinauto-recorder"),
	("", "python -m pywinauto_recorder"),
])
def test_program_name_in_the_help(monkeypatch, argv0, program_name):
	"""Tests the name of the program displayed by the help of the command line."""
	monkeypatch.setattr(sys, "argv", [argv0])
	monkeypatch.setattr(app.os.path, "basename", lambda path: path.replace("/", "\\").split("\\")[-1])
	assert app._program_name() == program_name


@pytest.mark.parametrize("ac_line_status, on_battery", [(0, True), (1, False), (255, False), (None, False)])
def test_is_on_battery(ac_line_status, on_battery):
	"""Tests that only the status 'on battery' (0) is a reason to warn, not the unknown status (255)."""
	assert app.is_on_battery(ac_line_status) is on_battery


def _fake_windll(get_system_power_status=None, console_window=0, console_process_ids=(), message_boxes=None,
                 shown_windows=None):
	"""Returns a fake ctypes.windll with the functions used by the application."""
	def get_console_process_list(process_ids, size):
		for i, process_id in enumerate(console_process_ids[:size]):
			process_ids[i] = process_id
		return len(console_process_ids)

	return types.SimpleNamespace(
		kernel32=types.SimpleNamespace(
			GetSystemPowerStatus=get_system_power_status,
			GetConsoleWindow=lambda: console_window,
			GetConsoleProcessList=get_console_process_list),
		user32=types.SimpleNamespace(
			MessageBoxW=lambda *args: message_boxes.append(args),
			ShowWindow=lambda *args: shown_windows.append(args)))


@pytest.mark.parametrize("ac_line_status", [0, 1, 255])
def test_get_ac_line_status(monkeypatch, ac_line_status):
	"""Tests that the AC power status is read as an unsigned byte: the unknown status is 255, not -1."""
	def get_system_power_status(status_pointer):
		status_pointer.contents.ACLineStatus = ac_line_status
		return 1

	monkeypatch.setattr(app, "windll", _fake_windll(get_system_power_status))
	assert app.get_ac_line_status() == ac_line_status


def test_get_ac_line_status_when_windows_cannot_give_it(monkeypatch):
	"""Tests that the AC power status is None when GetSystemPowerStatus fails."""
	monkeypatch.setattr(app, "windll", _fake_windll(lambda status_pointer: 0))
	assert app.get_ac_line_status() is None


@pytest.mark.parametrize("ac_line_status, nb_message_boxes", [(0, 1), (1, 0), (255, 0), (None, 0)])
def test_warn_if_on_battery(monkeypatch, ac_line_status, nb_message_boxes):
	"""Tests that the modal message box is displayed only when the computer runs on battery."""
	message_boxes = []
	monkeypatch.setattr(app, "windll", _fake_windll(message_boxes=message_boxes))
	monkeypatch.setattr(app, "get_ac_line_status", lambda: ac_line_status)
	app.warn_if_on_battery()
	assert len(message_boxes) == nb_message_boxes


@pytest.mark.parametrize("other_process_names, minimize", [
	([], True),  # double click on pywinauto_recorder.exe
	(["pywinauto-recorder.exe"], True),  # launcher of the command installed by pip, started from a shortcut
	(["cmd.exe"], False),
	(["PowerShell.exe", "pywinauto-recorder.exe"], False),
	(["pwsh.exe"], False),
])
def test_should_minimize_console(other_process_names, minimize):
	"""Tests that the console is minimized only when no command interpreter is attached to it."""
	assert app.should_minimize_console(other_process_names) is minimize


@pytest.mark.parametrize("console_window, other_process_names, minimized", [
	(1234, [], True), (1234, ["cmd.exe"], False), (0, [], False)])
def test_minimize_console_if_owned(monkeypatch, console_window, other_process_names, minimized):
	"""Tests that the console is minimized only if it exists and belongs to Pywinauto Recorder."""
	shown_windows = []
	monkeypatch.setattr(app, "windll", _fake_windll(console_window=console_window, shown_windows=shown_windows))
	monkeypatch.setattr(app, "_other_console_process_names", lambda: other_process_names)
	app.minimize_console_if_owned()
	assert shown_windows == ([(console_window, app.win32con.SW_MINIMIZE)] if minimized else [])


def test_console_process_ids(monkeypatch):
	"""Tests the IDs of the processes attached to the console, also when there are more than the first array can hold."""
	process_ids = tuple(range(100, 140))
	monkeypatch.setattr(app, "windll", _fake_windll(console_process_ids=process_ids))
	assert app._console_process_ids() == list(process_ids)


class _FakeRecorderForReplay:
	"""Replaces the Recorder in the tests of the replay of a file."""
	instances = []

	def __init__(self):
		self.played_back = []
		self.result = True
		_FakeRecorderForReplay.instances.append(self)

	def playback(self, str_code='', filename=''):
		self.played_back.append(filename)
		return _FakeRecorderForReplay.result


@pytest.fixture
def fake_recorder_for_replay(monkeypatch):
	"""Replaces the Recorder of the application, and makes stdin a non interactive input."""
	_FakeRecorderForReplay.instances = []
	_FakeRecorderForReplay.result = True
	monkeypatch.setattr(app, "Recorder", _FakeRecorderForReplay)
	monkeypatch.setattr(sys, "stdin", None)
	return _FakeRecorderForReplay


@pytest.mark.parametrize("result, exit_code", [(True, app.EXIT_OK), (False, app.EXIT_REPLAY_FAILED)])
def test_replay_file(tmp_path, fake_recorder_for_replay, result, exit_code):
	"""Tests the exit code of the replay of a file, when the script succeeds or fails."""
	script = tmp_path / "recorded.py"
	script.write_text("pass\n")
	fake_recorder_for_replay.result = result
	assert app.replay_file(str(script)) == exit_code
	assert [recorder_.played_back for recorder_ in fake_recorder_for_replay.instances] == [[str(script)]]


def test_replay_file_not_found(tmp_path, fake_recorder_for_replay):
	"""Tests that a missing file gives EXIT_FILE_NOT_FOUND at once, without starting a recorder."""
	assert app.replay_file(str(tmp_path / "missing.py")) == app.EXIT_FILE_NOT_FOUND
	assert fake_recorder_for_replay.instances == []


@pytest.mark.parametrize("argv, expected", [(["script.py"], ("replay", "script.py")), ([], ("interactive", ""))])
def test_main(monkeypatch, argv, expected):
	"""Tests that main() warns if the computer runs on battery, in both modes, then replays or starts the recorder."""
	calls = []
	monkeypatch.setattr(app, "warn_if_on_battery", lambda: calls.append("warn"))
	monkeypatch.setattr(app, "minimize_console_if_owned", lambda: calls.append("minimize"))
	monkeypatch.setattr(app, "replay_file", lambda filename: calls.append(("replay", filename)) or 7)
	monkeypatch.setattr(app, "run_interactive", lambda args: calls.append(("interactive", args.filename)) or 8)
	assert app.main(argv) == (7 if argv else 8)
	assert calls == ["warn", "minimize", expected]


@pytest.mark.skipif(sys.platform != "win32", reason="The package imports Windows modules")
@pytest.mark.parametrize("command", [
	[sys.executable, "-m", "pywinauto_recorder", "--help"],
	[sys.executable, str(PACKAGE_DIR.parent / "pywinauto_recorder.py"), "--help"],
])
def test_help_of_the_command_line(command):
	"""Tests 'python -m pywinauto_recorder --help' and the launcher, in a new process."""
	completed = subprocess.run(command, capture_output=True, text=True, timeout=120, cwd=str(PACKAGE_DIR.parent))
	assert completed.returncode == 0, completed.stderr
	assert "--no_splash_screen" in completed.stdout


################################################################################################
#                              Recorder: replay and keyboard shortcuts                         #
################################################################################################

class _FakeOverlay:
	def refresh(self):
		pass


@pytest.mark.parametrize("playback_result", [True, False])
def test_recorder_playback_returns_the_result_and_stops(monkeypatch, playback_result):
	"""Tests that Recorder.playback returns the result of the replay and sets the mode back to STOP."""
	modes = []
	fake_recorder = types.SimpleNamespace(main_overlay=_FakeOverlay(), mode=None)
	monkeypatch.setattr(recorder, "playback", lambda str_code, filename: modes.append(fake_recorder.mode) or playback_result)
	assert recorder.Recorder.playback(fake_recorder, filename="script.py") is playback_result
	assert modes == [recorder.RecorderMode.PLAY]
	assert fake_recorder.mode is recorder.RecorderMode.STOP


def test_recorder_playback_stops_when_the_replay_is_interrupted(monkeypatch):
	"""Tests that the mode is set back to STOP even if the replay is interrupted (KeyboardInterrupt)."""
	def interrupted_playback(str_code, filename):
		raise KeyboardInterrupt

	fake_recorder = types.SimpleNamespace(main_overlay=_FakeOverlay(), mode=None)
	monkeypatch.setattr(recorder, "playback", interrupted_playback)
	with pytest.raises(KeyboardInterrupt):
		recorder.Recorder.playback(fake_recorder, str_code="pass")
	assert fake_recorder.mode is recorder.RecorderMode.STOP


class _FakeKeyboard:
	"""Replaces the 'keyboard' module: CTRL and ALT (or CTRL and SHIFT) are pressed."""
	SCAN_CODES = {"ctrl": 29, "alt": 56, "shift": 42}

	def __init__(self, *pressed):
		self._pressed_events = {self.SCAN_CODES[name]: True for name in pressed}

	def key_to_scan_codes(self, name):
		return (self.SCAN_CODES[name],)


def _fake_recorder_for_shortcuts(mode):
	"""Returns an object that records the actions started by the keyboard shortcuts."""
	actions = []
	return types.SimpleNamespace(
		mode=mode, event_list=[], actions=actions, smart_mode=False,
		_Recorder__start_stop_recording_by_key=lambda: actions.append("record"),
		_Recorder__display_found_element_by_key=lambda: actions.append("copy code"),
		_Recorder__start_stop_displaying_info_by_key=lambda: actions.append("info"))


@pytest.mark.parametrize("key, pressed, action", [
	("r", ("ctrl", "alt"), "record"), ("D", ("ctrl", "shift"), "info"), ("F", ("ctrl", "shift"), "copy code")])
def test_shortcuts_are_ignored_during_a_replay(monkeypatch, key, pressed, action):
	"""Tests that the keyboard shortcuts work when the recorder is stopped, and are ignored during a replay."""
	monkeypatch.setattr(recorder, "keyboard", _FakeKeyboard(*pressed))
	event = types.SimpleNamespace(name=key, event_type="up")
	stopped = _fake_recorder_for_shortcuts(recorder.RecorderMode.STOP)
	recorder.Recorder._Recorder__key_on(stopped, event)
	assert stopped.actions == [action]
	playing = _fake_recorder_for_shortcuts(recorder.RecorderMode.PLAY)
	recorder.Recorder._Recorder__key_on(playing, event)
	assert playing.actions == []
	assert playing.event_list == []
