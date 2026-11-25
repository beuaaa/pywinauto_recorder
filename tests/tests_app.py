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
		self.calls = []
		_FakeRecorderForReplay.instances.append(self)

	def playback(self, str_code='', filename=''):
		self.played_back.append(filename)
		self.calls.append("playback")
		if isinstance(_FakeRecorderForReplay.result, BaseException):
			raise _FakeRecorderForReplay.result
		return _FakeRecorderForReplay.result

	def quit(self):
		self.calls.append("quit")


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
	assert fake_recorder_for_replay.instances[0].calls == ["playback", "quit"]


def test_replay_file_quits_the_recorder_after_an_interruption(tmp_path, fake_recorder_for_replay):
	"""Tests that the recorder quits, so that the application exits, even if the replay is interrupted by CTRL+C."""
	script = tmp_path / "recorded.py"
	script.write_text("pass\n")
	fake_recorder_for_replay.result = KeyboardInterrupt()
	with pytest.raises(KeyboardInterrupt):
		app.replay_file(str(script))
	assert fake_recorder_for_replay.instances[0].calls == ["playback", "quit"]


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


_REPLAY_LAUNCHER = """
import faulthandler, runpy, sys
faulthandler.dump_traceback_later(60, exit=True)  # a blocked replay prints the stacks of its threads, then exits
sys.argv = ["pywinauto-recorder", sys.argv[1]]
runpy.run_module("pywinauto_recorder", run_name="__main__", alter_sys=True)
"""


def _run_in_a_new_process(command, timeout, cwd):
	"""
	Runs a command in a new process, and kills at the end the processes it started.

	The output is written in a file, not in a pipe that a remaining child process could keep open, so that a blocked
	process cannot block the tests.

	:return: The exit code, or None if the process did not end within the timeout, the output, and the number of the
		child processes still running when the process ended
	"""
	import psutil
	import tempfile
	import time
	with tempfile.TemporaryFile() as output:
		process = subprocess.Popen(command, stdout=output, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, cwd=cwd)
		children = set()
		deadline = time.monotonic() + timeout
		while process.poll() is None and time.monotonic() < deadline:
			try:
				children.update(psutil.Process(process.pid).children(recursive=True))
			except psutil.Error:
				pass
			time.sleep(0.2)
		exit_code = process.poll()
		if exit_code is None:
			process.kill()
			process.wait()
		remaining_children = 0
		for child in children:
			try:
				if child.is_running():
					remaining_children += 1
					child.kill()
			except psutil.Error:
				pass
		output.seek(0)
		return exit_code, output.read().decode(errors="replace"), remaining_children


@pytest.mark.skipif(sys.platform != "win32", reason="The recorder needs Windows")
@pytest.mark.parametrize("script, expected_exit_code", [
	("print('The script is replayed')\n", app.EXIT_OK),
	("raise RuntimeError('The script fails')\n", app.EXIT_REPLAY_FAILED),
])
def test_replay_of_a_file_in_a_new_process(tmp_path, script, expected_exit_code):
	"""
	Tests the replay of a file by the command line in a new process: the process ends with the exit code of the
	replay, and the process of the element information tooltip ends too. Before, the process of the tooltip kept the
	application running after the replay.
	"""
	recorded_file = tmp_path / "recorded.py"
	recorded_file.write_text(script)
	exit_code, output, remaining_children = _run_in_a_new_process(
		[sys.executable, "-c", _REPLAY_LAUNCHER, str(recorded_file)], timeout=120, cwd=str(PACKAGE_DIR.parent))
	diagnostic = "AC line status: " + str(app.get_ac_line_status()) + "\n" + output
	assert exit_code == expected_exit_code, diagnostic
	assert "Timeout (" not in output, diagnostic  # no stacks printed by faulthandler
	assert "Replaying: " + str(recorded_file) in output, diagnostic
	assert "did not stop" not in output, diagnostic  # the loop of the recorder ended at once
	assert remaining_children == 0, diagnostic


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


class _FakeKeyboardHooks:
	"""
	Replaces the keyboard hooks and keybd_event. The releases of the Shift key whose numbers are in received_releases
	are received by the hook, after delay seconds, by another thread like the thread of the 'keyboard' library.
	"""
	def __init__(self, received_releases, delay=0.0):
		self.received_releases = received_releases
		self.delay = delay
		self.callbacks = []
		self.unhooked = []
		self.releases = []

	def hook(self, callback):
		self.callbacks.append(callback)
		return "hook " + str(len(self.callbacks))

	def unhook(self, hooked):
		self.unhooked.append(hooked)

	def keybd_event(self, virtual_key, scan_code, flags, extra_info):
		self.releases.append(virtual_key)
		if len(self.releases) in self.received_releases:
			threading.Timer(self.delay, self.callbacks[-1], args=("left shift up",)).start()


def _check_keyboard_hooks(monkeypatch, fake_hooks, timeout, interval):
	"""Returns the result of keyboard_hooks_work with the fake keyboard hooks, and its duration in seconds."""
	import time
	monkeypatch.setattr(recorder, "keyboard", fake_hooks)
	monkeypatch.setattr(recorder, "win32api", types.SimpleNamespace(keybd_event=fake_hooks.keybd_event))
	start = time.monotonic()
	result = recorder.keyboard_hooks_work(timeout=timeout, interval=interval)
	return result, time.monotonic() - start


def test_keyboard_hooks_work_at_once(monkeypatch):
	"""Tests that the check returns as soon as the hook receives the release of the left Shift key."""
	fake_hooks = _FakeKeyboardHooks(received_releases={1}, delay=0.05)
	result, duration = _check_keyboard_hooks(monkeypatch, fake_hooks, timeout=5.0, interval=2.0)
	assert result is True
	assert duration < 1.0
	assert fake_hooks.releases == [0xA0]  # VK_LSHIFT
	assert fake_hooks.unhooked == ["hook 1"]


def test_keyboard_hooks_installed_late(monkeypatch):
	"""
	Tests that the release of the Shift key is sent again until the hook receives it: the hook is installed by a thread
	of the 'keyboard' library, which may not be ready yet. Before, a single release was sent, and the result was read
	at once.
	"""
	fake_hooks = _FakeKeyboardHooks(received_releases={3})
	result, _ = _check_keyboard_hooks(monkeypatch, fake_hooks, timeout=5.0, interval=0.1)
	assert result is True
	assert len(fake_hooks.releases) >= 3
	assert fake_hooks.unhooked == ["hook 1"]


def test_keyboard_hooks_do_not_work(monkeypatch):
	"""Tests that the check fails after the timeout if the hook receives nothing, and removes its hook."""
	fake_hooks = _FakeKeyboardHooks(received_releases=set())
	result, duration = _check_keyboard_hooks(monkeypatch, fake_hooks, timeout=0.3, interval=0.05)
	assert result is False
	assert duration >= 0.3
	assert len(fake_hooks.releases) >= 2
	assert fake_hooks.unhooked == ["hook 1"]
	assert recorder.KEYBOARD_HOOKS_CHECK_TIMEOUT == 1.0


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


################################################################################################
#                                     Tray icon and its menu                                   #
################################################################################################

import threading


class _FakeRecorderForMenu:
	"""Replaces the Recorder in the tests of the menu: it records the calls of its methods."""
	def __init__(self, mode=recorder.RecorderMode.STOP):
		self.mode = mode
		self.smart_mode = False
		self.relative_coordinate_mode = False
		self.process_menu_click_mode = False
		self.calls = []

	def start_recording(self):
		self.calls.append("start_recording")
		self.mode = recorder.RecorderMode.RECORD

	def stop_recording(self):
		self.calls.append("stop_recording")
		self.mode = recorder.RecorderMode.STOP

	def toggle_element_info(self):
		self.calls.append("toggle_element_info")


class _FakeClipboardReplay:
	"""Replaces the ClipboardReplay in the tests of the menu."""
	def __init__(self, running=False):
		self.running = running
		self.started = []

	def is_running(self):
		return self.running

	def start(self, code):
		self.started.append(code)
		return True


def _menu(recorder_=None, replay=None):
	"""Returns the tray menu with fake recorder and replay, and the TrayIcons."""
	icons = app.TrayIcons()
	return app.build_menu(recorder_ or _FakeRecorderForMenu(), replay or _FakeClipboardReplay(), icons), icons


def _entries(menu):
	"""Returns the entries of a menu by text, without the shortcut: {"Start recording": MenuItem, ...}."""
	return {item.current_text().split("\t")[0]: item for item in menu if item is not app.MENU_SEPARATOR}


def test_menu_item_values_and_functions():
	"""Tests that the text, icon and state of an entry can be values or functions evaluated when the menu is displayed."""
	state = {"enabled": False}
	item = app.MenuItem(lambda: "Text", "icon.ico", enabled=lambda: state["enabled"])
	assert (item.current_text(), item.current_icon(), item.is_enabled()) == ("Text", "icon.ico", False)
	state["enabled"] = True
	assert item.is_enabled() is True
	assert app.MenuItem("Text").is_enabled() is True


def test_menu_when_the_recorder_is_stopped():
	"""Tests the texts and icons of the menu when the recorder is stopped."""
	menu, icons = _menu()
	assert [item.current_text() if item is not app.MENU_SEPARATOR else None for item in menu] == [
		"Start recording\t\t" + recorder.HOTKEY_RECORD,
		"Start replaying clipboard",
		"Start displaying element information\t" + recorder.HOTKEY_ELEMENT_INFO,
		"Start Smart mode\t\t" + recorder.HOTKEY_SMART_MODE,
		None, "Open output folder", "Process events", None, "Help", "Quit"]
	assert [item.current_icon() for item in menu[:4]] == [icons.record, icons.play, icons.search, icons.light_on]
	assert all(item.is_enabled() for item in menu if item is not app.MENU_SEPARATOR)


def test_menu_depends_on_the_state_of_the_recorder():
	"""Tests the texts and icons of the menu during a recording, the display of information and the Smart mode."""
	fake_recorder = _FakeRecorderForMenu(recorder.RecorderMode.RECORD)
	fake_recorder.smart_mode = True
	menu, icons = _menu(fake_recorder)
	entries = _entries(menu)
	assert entries["Stop recording"].current_icon() == icons.stop
	assert entries["Stop Smart mode"].current_icon() == icons.stop
	fake_recorder.mode = recorder.RecorderMode.INFO
	entries = _entries(menu)  # the same menu, displayed again
	assert entries["Stop displaying element information"].current_icon() == icons.stop
	assert "Start recording" in entries


def test_menu_during_a_replay_of_the_clipboard():
	"""
	Tests that, during a replay of the clipboard, the entries that would interfere with the replay are grayed,
	"Quit" included, and that the others stay enabled.
	"""
	menu, icons = _menu(replay=_FakeClipboardReplay(running=True))
	entries = _entries(menu)
	assert entries["Wait end of replay"].current_icon() == icons.stop
	grayed = {text for text, item in entries.items() if not item.is_enabled()}
	assert grayed == {"Start recording", "Wait end of replay", "Start displaying element information",
	                  "Start Smart mode", "Quit"}
	help_entries = _entries(entries["Help"].submenu)
	assert not help_entries["About"].is_enabled()
	assert help_entries["Web site"].is_enabled()
	assert all(item.is_enabled() for item in entries["Process events"].submenu[:2])


def test_menu_set_text_and_set_combobox_are_not_implemented_yet():
	"""Tests that set_text and set_combobox are grayed entries that tell they are not implemented yet."""
	menu, _ = _menu()
	process_events = _entries(_entries(menu)["Process events"].submenu)
	for name in ("set_text", "set_combobox"):
		item = process_events[name + " (not implemented yet)"]
		assert not item.is_enabled()
		assert item.action is None


def test_menu_process_events_toggles():
	"""Tests the entries of 'Process events' that switch the processing of the recorded events."""
	fake_recorder = _FakeRecorderForMenu()
	menu, icons = _menu(fake_recorder)
	process_events = _entries(menu)["Process events"].submenu
	relative_coordinates, menu_click = process_events[0], process_events[1]
	assert relative_coordinates.current_icon() == icons.cross
	relative_coordinates.action(None)
	assert fake_recorder.relative_coordinate_mode is True
	assert relative_coordinates.current_icon() == icons.check
	menu_click.action(None)
	assert fake_recorder.process_menu_click_mode is True
	assert menu_click.current_icon() == icons.check


def test_menu_actions(monkeypatch, tmp_path):
	"""Tests the actions of the entries of the menu."""
	fake_recorder = _FakeRecorderForMenu()
	replay = _FakeClipboardReplay()
	menu, _ = _menu(fake_recorder, replay)
	entries = _entries(menu)
	entries["Start recording"].action(None)
	assert fake_recorder.calls == ["start_recording"]
	_entries(menu)["Stop recording"].action(None)
	assert fake_recorder.calls == ["start_recording", "stop_recording"]

	monkeypatch.setattr(app.pyperclip, "paste", lambda: "print('replayed')")
	entries["Start replaying clipboard"].action(None)
	assert replay.started == ["print('replayed')"]

	entries["Start displaying element information"].action(None)
	assert fake_recorder.calls[-1] == "toggle_element_info"

	entries["Start Smart mode"].action(None)
	assert fake_recorder.smart_mode is True

	quit_calls = []
	entries["Quit"].action(types.SimpleNamespace(quit=lambda: quit_calls.append("quit")))
	assert quit_calls == ["quit"]


def test_open_output_folder(monkeypatch, tmp_path):
	"""Tests that 'Open output folder' creates the folder if needed and opens it."""
	opened = []
	output_folder = tmp_path / "Pywinauto recorder"
	monkeypatch.setattr(app, "OUTPUT_FOLDER", output_folder)
	monkeypatch.setattr(app.os, "startfile", lambda path: opened.append(path), raising=False)
	app.open_output_folder()
	assert output_folder.is_dir()
	assert opened == [str(output_folder)]


def test_open_web_site(monkeypatch):
	"""Tests that 'Web site' opens the documentation in the default browser."""
	opened = []
	monkeypatch.setattr(app.webbrowser, "open", lambda url: opened.append(url))
	app.open_web_site()
	assert opened == ["https://pywinauto-recorder.readthedocs.io"]


@pytest.mark.parametrize("mode, calls", [
	(recorder.RecorderMode.RECORD, ["stop_recording"]), (recorder.RecorderMode.INFO, []), (recorder.RecorderMode.STOP, [])])
def test_show_about_stops_the_recording_first(monkeypatch, mode, calls):
	"""Tests that 'About' writes the recording in progress (stop_recording) before displaying the splash screen."""
	splash_screens = []
	monkeypatch.setattr(app, "display_splash_screen", lambda *args, **kwargs: splash_screens.append(args))
	fake_recorder = _FakeRecorderForMenu(mode)
	app.show_about(fake_recorder)
	assert fake_recorder.calls == calls
	assert fake_recorder.mode is recorder.RecorderMode.STOP
	assert len(splash_screens) == 1


class _FakeRecorderForQuit:
	"""Replaces the Recorder in the tests of Recorder.quit: it records the calls in their order."""
	def __init__(self, mode, alive_after_join=False):
		self.mode = mode
		self.calls = []
		self.alive_after_join = alive_after_join
		self.main_overlay = types.SimpleNamespace(
			clear_all=lambda: self.calls.append("clear overlay"), refresh=lambda: self.calls.append("refresh overlay"))
		self.element_info_tooltip = "tooltip"

	def stop_recording(self):
		self.calls.append("stop_recording")
		self.mode = recorder.RecorderMode.STOP

	def join(self, timeout=None):
		self.calls.append(("join", timeout, self.mode))

	def is_alive(self):
		return self.alive_after_join


@pytest.mark.parametrize("mode, calls", [
	(recorder.RecorderMode.RECORD, ["stop_recording"]), (recorder.RecorderMode.INFO, []), (recorder.RecorderMode.STOP, [])])
def test_recorder_quit(capsys, mode, calls):
	"""
	Tests that quit() writes the recording in progress, stops the loop of the recorder and waits for its end at most
	QUIT_TIMEOUT seconds, and then stops the element information tooltip. Before, the recording in progress was lost.
	"""
	fake_recorder = _FakeRecorderForQuit(mode)
	recorder.Recorder.quit(fake_recorder)
	assert fake_recorder.calls == calls + [
		"clear overlay", "refresh overlay", ("join", recorder.QUIT_TIMEOUT, recorder.RecorderMode.QUIT)]
	assert not hasattr(fake_recorder, "element_info_tooltip")
	assert capsys.readouterr().out == "Quit\n"


def test_recorder_quit_does_not_wait_forever(capsys):
	"""Tests that quit() returns, and stops the tooltip, when the loop of the recorder is blocked."""
	fake_recorder = _FakeRecorderForQuit(recorder.RecorderMode.INFO, alive_after_join=True)
	recorder.Recorder.quit(fake_recorder)
	assert not hasattr(fake_recorder, "element_info_tooltip")
	assert recorder.QUIT_TIMEOUT == 10
	assert capsys.readouterr().out == "The recorder did not stop within 10 seconds.\nQuit\n"


@pytest.mark.parametrize("mode, calls, new_mode", [
	(recorder.RecorderMode.STOP, [], recorder.RecorderMode.INFO),
	(recorder.RecorderMode.INFO, [], recorder.RecorderMode.STOP),
	(recorder.RecorderMode.RECORD, ["stop_recording"], recorder.RecorderMode.INFO),
])
def test_recorder_toggle_element_info(mode, calls, new_mode):
	"""
	Tests that displaying element information stops the recording in progress with stop_recording(), so that the
	recorded script is written. Before, the recording was lost (menu and CTRL+SHIFT+D).
	"""
	fake_recorder = _FakeRecorderForMenu(mode)
	recorder.Recorder.toggle_element_info(fake_recorder)
	assert fake_recorder.calls == calls
	assert fake_recorder.mode is new_mode


class _BlockingRecorder:
	"""A recorder whose replay lasts until 'finish' is set."""
	def __init__(self):
		self.finish = threading.Event()
		self.codes = []

	def playback(self, str_code='', filename=''):
		self.codes.append(str_code)
		self.finish.wait(5)
		return True


def test_clipboard_replay_runs_in_a_thread():
	"""
	Tests that the clipboard is replayed in a thread: start() returns at once, is_running() is True until the end of
	the replay, and a second replay cannot start meanwhile.
	"""
	blocking_recorder = _BlockingRecorder()
	replay = app.ClipboardReplay(blocking_recorder)
	assert replay.is_running() is False
	assert replay.start("first code") is True
	assert replay.is_running() is True
	assert replay.start("second code") is False
	blocking_recorder.finish.set()
	replay._thread.join(5)
	assert replay.is_running() is False
	assert blocking_recorder.codes == ["first code"]
	assert replay.start("third code") is True
	replay._thread.join(5)
	assert blocking_recorder.codes == ["first code", "third code"]


class _FakeWin32Gui:
	"""Records the native menus built by SysTrayIcon._fill_menu."""
	def __init__(self):
		self.menus = {}

	def CreatePopupMenu(self):
		handle = 1000 + len(self.menus)
		self.menus[handle] = []
		return handle

	def InsertMenu(self, menu, position, flags, item_id, text):
		self.menus[menu].insert(position, "separator")

	def InsertMenuItem(self, menu, position, by_position, menu_item_info):
		self.menus[menu].insert(position, menu_item_info)

	@staticmethod
	def LOWORD(value):
		return value & 0xFFFF


def test_tray_menu_is_built_from_the_entries(monkeypatch):
	"""Tests that the native menu has the current texts, the grayed state, the submenus, and calls the right actions."""
	fake_win32gui = _FakeWin32Gui()
	monkeypatch.setattr(app, "win32gui", fake_win32gui)
	monkeypatch.setattr(app.win32gui_struct, "PackMENUITEMINFO", lambda **kwargs: (kwargs, None))
	tray = object.__new__(app.SysTrayIcon)
	tray._actions_by_id = {}
	monkeypatch.setattr(tray, "_menu_bitmap", lambda icon: "bitmap of " + icon, raising=False)
	clicked = []
	items = [
		app.MenuItem(lambda: "First", "first.ico", action=lambda tray_: clicked.append("first")),
		app.MENU_SEPARATOR,
		app.MenuItem("Submenu", None, submenu=[app.MenuItem("Grayed", enabled=False)]),
		app.MenuItem("Last", action=lambda tray_: clicked.append("last")),
	]
	menu = fake_win32gui.CreatePopupMenu()
	tray._fill_menu(menu, items)
	entries = fake_win32gui.menus[menu]
	assert [entry if entry == "separator" else entry["text"] for entry in entries] == ["First", "separator", "Submenu", "Last"]
	assert entries[0]["hbmpItem"] == "bitmap of first.ico" and entries[0]["fState"] is None
	submenu_entries = fake_win32gui.menus[entries[2]["hSubMenu"]]
	assert submenu_entries[0]["text"] == "Grayed"
	assert submenu_entries[0]["fState"] == app.win32con.MFS_GRAYED
	tray.command(None, None, entries[3]["wID"], None)
	tray.command(None, None, entries[0]["wID"], None)
	tray.command(None, None, submenu_entries[0]["wID"], None)  # a grayed entry has no action
	assert clicked == ["last", "first"]


def test_tray_icon_removed_even_if_on_quit_fails(monkeypatch):
	"""
	Tests that the tray icon is removed and run() stopped, even if on_quit raises an exception. Before, the application
	kept running without its tray icon.
	"""
	calls = []
	monkeypatch.setattr(app, "win32gui", types.SimpleNamespace(
		NIM_DELETE=2, Shell_NotifyIcon=lambda message, nid: calls.append(("Shell_NotifyIcon", message, nid)),
		PostQuitMessage=lambda exit_code: calls.append(("PostQuitMessage", exit_code))))
	tray = object.__new__(app.SysTrayIcon)
	tray.hwnd = 42

	def on_quit(tray_):
		raise RuntimeError("The recorded script cannot be written")
	tray.on_quit = on_quit
	with pytest.raises(RuntimeError):
		tray.destroy(42, None, None, None)
	assert calls == [("Shell_NotifyIcon", 2, (42, 0)), ("PostQuitMessage", 0)]


################################################################################################
#                               Version check and splash screen                                #
################################################################################################

import socket
import ssl
import urllib.error

VERSION_INFO_RC = PACKAGE_DIR.parent / "bin" / "VersionInfo.rc"


@pytest.fixture
def version_not_checked():
	"""check_version() checks the version once: this fixture clears its cache before and after the test."""
	app.check_version.cache_clear()
	yield
	app.check_version.cache_clear()


@pytest.mark.parametrize("older, newer", [
	("0.6.8", "0.6.9"), ("0.6.9", "0.10.0"), ("0.6.9", "1.0.0.dev0"), ("1.0.0.dev0", "1.0.0a1"), ("1.0.0a1", "1.0.0b1"),
	("1.0.0b1", "1.0.0b2"), ("1.0.0b2", "1.0.0rc1"), ("1.0.0rc1", "1.0.0"), ("1.0.0", "1.0.1"), ("1.0.0", "1.0.0.post1"),
])
def test_version_key_order(older, newer):
	"""Tests that the versions are compared as numbers, and that a pre-release comes before its final release."""
	assert app.version_key(older) < app.version_key(newer)
	assert app.version_key(newer) > app.version_key(older)


@pytest.mark.parametrize("version, same_version", [
	("1.0", "1.0.0"), ("v0.6.9", "0.6.9"), (" 0.6.9\r\n", "0.6.9"), ("1.0.0B1", "1.0.0b1"), ("1.0.0-beta.1", "1.0.0b1"),
])
def test_version_key_same_version(version, same_version):
	assert app.version_key(version) == app.version_key(same_version)


@pytest.mark.parametrize("installed, latest, console_message, version_line", [
	("0.6.9", "0.6.9", "Your Pywinauto recorder is up to date.", "Pywinauto recorder 0.6.9"),
	("0.6.8", "0.6.9", "Pywinauto recorder v0.6.9 is available!",
	 "Your version of Pywinauto recorder (0.6.8) is not up to date (0.6.9 is available!)"),
	("1.0.0.dev0", "0.6.9", "Your Pywinauto recorder 1.0.0.dev0 is newer than the latest published version (0.6.9).",
	 "Pywinauto recorder 1.0.0.dev0"),
	("0.6.9", None, "Cannot check whether a newer version of Pywinauto recorder is available.",
	 "Pywinauto recorder 0.6.9"),
])
def test_version_messages(installed, latest, console_message, version_line):
	"""
	Tests the messages of the console and of the splash screen. Before, a development version was "not up to date"
	because it was not the latest published version, and "up to date" when the version could not be checked.
	"""
	assert app.version_messages(installed, latest) == (console_message, version_line)


class _FakeResponse:
	"""Replaces the response of urllib.request.urlopen."""
	def __init__(self, content):
		self.content = content

	def __enter__(self):
		return self

	def __exit__(self, *exc_info):
		return False

	def read(self):
		return self.content


def test_get_latest_version_reads_version_info_rc(monkeypatch):
	"""Tests that the latest version is read in bin/VersionInfo.rc of the branch master, with a timeout."""
	if not VERSION_INFO_RC.is_file():
		pytest.skip("bin/VersionInfo.rc is not next to the package")
	requests = []

	def urlopen(url, timeout=None):
		requests.append((url, timeout))
		return _FakeResponse(VERSION_INFO_RC.read_bytes())
	monkeypatch.setattr(app.urllib.request, "urlopen", urlopen)
	assert app.get_latest_version() == pywinauto_recorder.__version__
	assert requests == [(app.LATEST_VERSION_URL, app.VERSION_CHECK_TIMEOUT)]
	assert app.LATEST_VERSION_URL.endswith("/beuaaa/pywinauto_recorder/master/bin/VersionInfo.rc")
	assert 0 < app.VERSION_CHECK_TIMEOUT <= 10


@pytest.mark.parametrize("content", [
	b"", b"<html>Proxy authentication required</html>", b'VALUE "ProductVersion", ""',
	b'VALUE "ProductVersion", "unknown"', b"\xff\xfe\x00"])
def test_get_latest_version_unexpected_content(monkeypatch, content):
	"""Tests that an unexpected content, a page of a proxy for example, gives no version."""
	monkeypatch.setattr(app.urllib.request, "urlopen", lambda url, timeout=None: _FakeResponse(content))
	assert app.get_latest_version() is None


@pytest.mark.parametrize("error", [
	urllib.error.URLError("no network"), socket.timeout("timed out"), ssl.SSLError("certificate verify failed"),
	ConnectionResetError("connection reset"), ValueError("unknown url type")])
def test_get_latest_version_network_errors(monkeypatch, error):
	"""
	Tests that the errors are ignored. Before, only URLError was caught: another error, a connection reset while
	reading the response for example, stopped Pywinauto Recorder at startup.
	"""
	def urlopen(url, timeout=None):
		raise error
	monkeypatch.setattr(app.urllib.request, "urlopen", urlopen)
	assert app.get_latest_version() is None


def test_check_version_is_done_once(monkeypatch, capsys, version_not_checked):
	"""
	Tests that check_version() prints the message in the console and returns the line of the splash screen, and that
	the version is checked only once: "About" displays the splash screen again without a new request.
	"""
	requests = []
	monkeypatch.setattr(app, "get_latest_version", lambda: requests.append("request") or "0.0.1")
	version_line = app.check_version()
	assert version_line == "Pywinauto recorder " + pywinauto_recorder.__version__
	assert capsys.readouterr().out == ("Your Pywinauto recorder " + pywinauto_recorder.__version__ +
	                                   " is newer than the latest published version (0.0.1).\n")
	assert app.check_version() == version_line
	assert capsys.readouterr().out == ""
	assert requests == ["request"]


def test_splash_screen_text_lines():
	"""Tests that the splash screen displays the keyboard shortcuts of the recorder, at the same places as before."""
	text_lines = app._splash_screen_text_lines("Pywinauto recorder 1.2.3")
	assert len(text_lines) == 24
	assert text_lines[6] == "Pywinauto recorder 1.2.3"
	assert text_lines[9] == 'Record / Stop                            Smart mode On / Off'
	assert text_lines[10] == 'CTRL+ALT+R                                      CTRL+ALT+S'
	assert text_lines[13] == '                                                       CTRL+SHIFT+D'
	assert text_lines[16] == 'CTRL+SHIFT+F                                tray menu to quit'
	assert text_lines[23] == ''


class _FakeSplashOverlay:
	"""Replaces overlay_arrows_and_more.Overlay: it keeps what is displayed at each refresh."""
	def __init__(self, transparency=0.0):
		self.items = []
		self.frames = []

	def clear_all(self):
		self.items = []

	def add(self, **kwargs):
		self.items.append(kwargs)

	def refresh(self):
		self.frames.append(list(self.items))


def _display_splash_screen(monkeypatch, cursor_positions):
	"""
	Displays the splash screen on a 1920x1080 screen, with fake overlays and the given positions of the mouse cursor.

	:return: The foreground and the background overlays
	"""
	overlays = []

	def make_overlay(transparency=0.0):
		overlays.append(_FakeSplashOverlay(transparency))
		return overlays[-1]
	positions = iter(cursor_positions)
	monkeypatch.setattr(app.oaam, "Overlay", make_overlay)
	monkeypatch.setattr(app, "time", types.SimpleNamespace(sleep=lambda seconds: None))
	monkeypatch.setattr(app, "GetSystemMetrics", lambda index: (1920, 1080)[index])
	monkeypatch.setattr(app, "GetCursorPos", lambda: next(positions))
	monkeypatch.setattr(app, "check_version", lambda: "Pywinauto recorder 1.2.3")
	app.display_splash_screen()
	return overlays


def _displayed_text_lines(frame):
	"""Returns the 24 lines of text of a frame of the splash screen: they are drawn before the icons."""
	return [item["text"] for item in frame[:24]]


IN_SPLASH_SCREEN = (960, 540)
OUT_OF_SPLASH_SCREEN = (100, 100)


def test_splash_screen_closed_when_the_mouse_cursor_moves_out_of_it(monkeypatch):
	"""Tests that the splash screen, displayed under the mouse cursor, is closed when the mouse cursor moves out of it."""
	foreground, background = _display_splash_screen(
		monkeypatch, [IN_SPLASH_SCREEN, IN_SPLASH_SCREEN, IN_SPLASH_SCREEN, OUT_OF_SPLASH_SCREEN])
	frames = [_displayed_text_lines(frame) for frame in foreground.frames[:-1]]
	assert len(frames) == 3
	assert all(text_lines[6] == "Pywinauto recorder 1.2.3" for text_lines in frames)
	assert [text_lines[23] for text_lines in frames] == [
		'', 'To continue: move the mouse cursor out of this window', '']  # the message blinks
	assert foreground.frames[-1] == [] and background.frames[-1] == []  # the splash screen is closed


def test_splash_screen_closed_when_the_mouse_cursor_moves_over_it(monkeypatch):
	"""Tests that the splash screen, displayed away from the mouse cursor, is closed when the mouse cursor moves over it."""
	foreground, background = _display_splash_screen(
		monkeypatch, [OUT_OF_SPLASH_SCREEN, OUT_OF_SPLASH_SCREEN, IN_SPLASH_SCREEN])
	frames = [_displayed_text_lines(frame) for frame in foreground.frames[:-1]]
	assert [text_lines[23] for text_lines in frames] == ['', 'To continue: move the mouse cursor over this window']
	assert foreground.frames[-1] == [] and background.frames[-1] == []


def test_splash_screen_is_centered(monkeypatch):
	monkeypatch.setattr(app, "GetSystemMetrics", lambda index: (1920, 1080)[index])
	assert app._splash_rectangle() == (640, 270, 1280, 810)


@pytest.mark.parametrize("argv, events", [
	([], ["recorder", "splash screen", "tray icon"]),
	(["--no_splash_screen"], ["recorder", "tray icon"]),
])
def test_run_interactive(monkeypatch, argv, events):
	"""
	Tests that the recorder starts with the splash screen, unless --no_splash_screen is given, then runs the tray icon.
	Its double click displays element information, except during a replay of the clipboard, and "Quit" stops the
	recorder.
	"""
	happened = []
	trays = []
	fake_recorder = _FakeRecorderForMenu()
	fake_recorder.quit = lambda: fake_recorder.calls.append("quit")

	class FakeSysTrayIcon:
		def __init__(self, icon, hover_text, menu, on_quit=None, default_action=None):
			self.menu, self.on_quit, self.default_action = menu, on_quit, default_action
			trays.append(self)

		def run(self):
			happened.append("tray icon")

	monkeypatch.setattr(app, "Recorder", lambda: happened.append("recorder") or fake_recorder)
	monkeypatch.setattr(app, "display_splash_screen", lambda: happened.append("splash screen"))
	monkeypatch.setattr(app, "SysTrayIcon", FakeSysTrayIcon)
	assert app.run_interactive(app.parse_arguments(argv)) == app.EXIT_OK
	assert happened == events
	tray = trays[0]
	tray.default_action(tray)
	assert fake_recorder.calls == ["toggle_element_info"]
	monkeypatch.setattr(app.ClipboardReplay, "is_running", lambda self: True)
	tray.default_action(tray)
	assert fake_recorder.calls == ["toggle_element_info"]
	tray.on_quit(tray)
	assert fake_recorder.calls == ["toggle_element_info", "quit"]


@pytest.mark.parametrize("alive", [False, True])
def test_element_info_tooltip_does_not_block_the_quit(alive):
	"""
	Tests that deleting the tooltip stops its process, waiting for its end at most TOOLTIP_STOP_TIMEOUT seconds, and
	terminates it if it does not stop. Before, the recorder waited for it without a time limit.
	"""
	from pywinauto_recorder import element_observer
	calls = []

	class FakeProcess:
		def join(self, timeout=None):
			calls.append(("join", timeout))

		def is_alive(self):
			return alive

		def terminate(self):
			calls.append("terminate")

	quit_event = threading.Event()
	tooltip = object.__new__(element_observer.ElementInfoTooltip)
	tooltip.info_mode_quit = quit_event
	tooltip.process = FakeProcess()
	del tooltip
	assert quit_event.is_set()
	assert calls == [("join", element_observer.TOOLTIP_STOP_TIMEOUT)] + (["terminate"] if alive else [])
	assert element_observer.TOOLTIP_STOP_TIMEOUT == 5


def test_element_info_tooltip_process_stopped_just_after_its_start(monkeypatch):
	"""
	Tests that the process of the tooltip ends when it is stopped before its loop starts, without calling quit() of its
	overlay: quit() of overlay_arrows_and_more never returns if the window of the overlay is not created yet.
	"""
	from pywinauto_recorder import element_observer
	overlays = []

	class FakeOverlay:
		def __init__(self, transparency=0.0):
			self.calls = []
			overlays.append(self)

		def clear_all(self):
			self.calls.append("clear_all")

		def refresh(self):
			self.calls.append("refresh")

		def quit(self):
			raise AssertionError("quit() never returns if the window of the overlay is not created yet")

	monkeypatch.setattr(element_observer.oaam, "Overlay", FakeOverlay)
	monkeypatch.setattr(element_observer, "Desktop", lambda **kwargs: None)
	monkeypatch.setattr(element_observer, "GetSystemMetrics", lambda index: 1920)
	quit_event = threading.Event()
	quit_event.set()
	element_observer.task(threading.Event(), quit_event)
	assert overlays[0].calls == ["clear_all", "refresh"]
