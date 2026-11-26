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
	for name in (
			"IconPyRec.ico", "record.ico", "stop.ico", "play.ico", "folder.ico", "search.ico", "light-on.ico",
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
	entry_points = [
		entry_point for entry_point in distribution.entry_points
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

# Bandit B404 and B603 (subprocess): the tests run Pywinauto Recorder in a new process, with the current Python
import subprocess  # nosec B404
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


def _fake_windll(
		get_system_power_status=None, console_window=0, console_process_ids=(), message_boxes=None, shown_windows=None):
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
	"""
	Replaces the Recorder in the tests of the replay of a file.

	playback() raises 'error' if it is not None, else it returns 'result'.
	"""

	instances = []
	result = True
	error = None

	def __init__(self):
		self.played_back = []
		self.calls = []
		_FakeRecorderForReplay.instances.append(self)

	def playback(self, str_code='', filename=''):
		self.played_back.append(filename)
		self.calls.append("playback")
		if _FakeRecorderForReplay.error is not None:
			raise _FakeRecorderForReplay.error
		return _FakeRecorderForReplay.result

	def quit(self):
		self.calls.append("quit")


@pytest.fixture
def fake_recorder_for_replay(monkeypatch):
	"""Replaces the Recorder of the application, and makes stdin a non interactive input."""
	_FakeRecorderForReplay.instances = []
	_FakeRecorderForReplay.result = True
	_FakeRecorderForReplay.error = None
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
	fake_recorder_for_replay.error = KeyboardInterrupt()
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
	completed = subprocess.run(  # nosec B603
		command, capture_output=True, text=True, timeout=120, cwd=str(PACKAGE_DIR.parent))
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
		process = subprocess.Popen(  # nosec B603
			command, stdout=output, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, cwd=cwd)
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
	Tests the replay of a file by the command line in a new process.

	The process ends with the exit code of the replay, and the process of the element information tooltip ends too.
	Before, the process of the tooltip kept the application running after the replay.
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
	Replaces the keyboard hooks and keybd_event.

	The releases of the Shift key whose numbers are in received_releases are received by the hook, after delay seconds,
	by another thread like the thread of the 'keyboard' library.
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
	Tests that the release of the Shift key is sent again until the hook receives it.

	The hook is installed by a thread of the 'keyboard' library, which may not be ready yet. Before, a single release
	was sent, and the result was read at once.
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
	Tests that, during a replay of the clipboard, the entries that would interfere with the replay are grayed.

	"Quit" is grayed too. The other entries stay enabled.
	"""
	menu, icons = _menu(replay=_FakeClipboardReplay(running=True))
	entries = _entries(menu)
	assert entries["Wait end of replay"].current_icon() == icons.stop
	grayed = {text for text, item in entries.items() if not item.is_enabled()}
	assert grayed == {
		"Start recording", "Wait end of replay", "Start displaying element information", "Start Smart mode", "Quit"}
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
	monkeypatch.setattr(app.os, "startfile", opened.append, raising=False)
	app.open_output_folder()
	assert output_folder.is_dir()
	assert opened == [str(output_folder)]


def test_open_web_site(monkeypatch):
	"""Tests that 'Web site' opens the documentation in the default browser."""
	opened = []
	monkeypatch.setattr(app.webbrowser, "open", opened.append)
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
	Tests that quit() writes the recording in progress, stops the recorder, and then stops the tooltip.

	It waits for the end of the loop of the recorder at most QUIT_TIMEOUT seconds, and then stops the element
	information tooltip. Before, the recording in progress was lost.
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
	Tests that displaying element information stops the recording in progress with stop_recording().

	The recorded script is written. Before, the recording was lost (menu and CTRL+SHIFT+D).
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
	Tests that the clipboard is replayed in a thread.

	start() returns at once, is_running() is True until the end of the replay, and a second replay cannot start
	meanwhile.
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


def test_tray_icon_message_handlers_return_an_int(monkeypatch):
	"""
	Tests that the handlers of the messages of the tray icon return an int: the result of the window procedure.

	Before, restart, destroy and command returned None, which recent versions of pywin32 refuse: "WNDPROC return value
	cannot be converted to LRESULT" was printed, twice when quitting for example.
	"""
	calls = []
	monkeypatch.setattr(app, "win32gui", types.SimpleNamespace(
		NIM_DELETE=2, LOWORD=lambda value: value & 0xFFFF,
		Shell_NotifyIcon=lambda message, nid: calls.append("Shell_NotifyIcon"),
		PostQuitMessage=lambda exit_code: calls.append("PostQuitMessage")))
	tray = object.__new__(app.SysTrayIcon)
	tray.hwnd = 42
	tray.on_quit = lambda tray_: calls.append("on_quit")
	tray.default_action = lambda tray_: calls.append("default_action")
	tray._actions_by_id = {1024: lambda tray_: calls.append("action")}
	tray.refresh_icon = lambda: calls.append("refresh_icon")
	tray.show_menu = lambda: calls.append("show_menu")
	results = [
		tray.restart(42, 0, 0, 0),
		tray.command(42, 0, 1024, 0),
		tray.command(42, 0, 9999, 0),  # an entry without action
		tray.notify(42, 0, 0, app.win32con.WM_LBUTTONDBLCLK),
		tray.notify(42, 0, 0, app.win32con.WM_RBUTTONUP),
		tray.destroy(42, 0, 0, 0)]
	assert [type(result) for result in results] == [int] * 6
	assert calls == [
		"refresh_icon", "action", "default_action", "show_menu", "on_quit", "Shell_NotifyIcon", "PostQuitMessage"]


def test_tray_icon_removed_even_if_on_quit_fails(monkeypatch):
	"""
	Tests that the tray icon is removed and run() stopped, even if on_quit raises an exception.

	Before, the application kept running without its tray icon.
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
	(
		"0.6.8", "0.6.9", "Pywinauto recorder v0.6.9 is available!",
		"Your version of Pywinauto recorder (0.6.8) is not up to date (0.6.9 is available!)"),
	(
		"1.0.0.dev0", "0.6.9", "Your Pywinauto recorder 1.0.0.dev0 is newer than the latest published version (0.6.9).",
		"Pywinauto recorder 1.0.0.dev0"),
	(
		"0.6.9", None, "Cannot check whether a newer version of Pywinauto recorder is available.",
		"Pywinauto recorder 0.6.9"),
])
def test_version_messages(installed, latest, console_message, version_line):
	"""
	Tests the messages of the console and of the splash screen.

	Before, a development version was "not up to date" because it was not the latest published version, and "up to
	date" when the version could not be checked.
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
	Tests that the errors are ignored.

	Before, only URLError was caught: another error, a connection reset while reading the response for example, stopped
	Pywinauto Recorder at startup.
	"""
	def urlopen(url, timeout=None):
		raise error
	monkeypatch.setattr(app.urllib.request, "urlopen", urlopen)
	assert app.get_latest_version() is None


def test_check_version_is_done_once(monkeypatch, capsys, version_not_checked):
	"""
	Tests that check_version() prints the message in the console and returns the line of the splash screen.

	The version is checked only once: "About" displays the splash screen again without a new request.
	"""
	requests = []
	monkeypatch.setattr(app, "get_latest_version", lambda: requests.append("request") or "0.0.1")
	version_line = app.check_version()
	assert version_line == "Pywinauto recorder " + pywinauto_recorder.__version__
	assert capsys.readouterr().out == (
		"Your Pywinauto recorder " + pywinauto_recorder.__version__ + " is newer than the latest published version (0.0.1).\n")
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
	"""Replaces overlay_arrows_and_more.Overlay: it keeps what is displayed at each refresh, and whether it is closed."""

	def __init__(self, transparency=0.0):
		self.items = []
		self.frames = []
		self.closed = False

	def quit(self, timeout=None):
		self.closed = True

	def clear_all(self):
		self.items = []

	def add(self, **kwargs):
		self.items.append(kwargs)

	def refresh(self):
		self.frames.append(list(self.items))


def _patch_splash_screen(monkeypatch, get_cursor_pos):
	"""
	Patches the splash screen: a 1920x1080 screen, fake overlays and the given function for the position of the mouse.

	:return: The list of the overlays that the splash screen will create: the foreground, then the background
	"""
	overlays = []

	def make_overlay(transparency=0.0):
		overlays.append(_FakeSplashOverlay(transparency))
		return overlays[-1]
	monkeypatch.setattr(app.oaam, "Overlay", make_overlay)
	monkeypatch.setattr(app, "time", types.SimpleNamespace(sleep=lambda seconds: None))
	monkeypatch.setattr(app, "GetSystemMetrics", lambda index: (1920, 1080)[index])
	monkeypatch.setattr(app, "GetCursorPos", get_cursor_pos)
	monkeypatch.setattr(app, "check_version", lambda: "Pywinauto recorder 1.2.3")
	return overlays


def _display_splash_screen(monkeypatch, cursor_positions):
	"""
	Displays the splash screen on a 1920x1080 screen, with fake overlays and the given positions of the mouse cursor.

	:return: The foreground and the background overlays
	"""
	positions = iter(cursor_positions)
	overlays = _patch_splash_screen(monkeypatch, lambda: next(positions))
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
	frames = [_displayed_text_lines(frame) for frame in foreground.frames]
	assert len(frames) == 3
	assert all(text_lines[6] == "Pywinauto recorder 1.2.3" for text_lines in frames)
	assert [text_lines[23] for text_lines in frames] == [
		'', 'To continue: move the mouse cursor out of this window', '']  # the message blinks
	assert foreground.closed and background.closed


def test_splash_screen_closed_when_the_mouse_cursor_moves_over_it(monkeypatch):
	"""Tests that the splash screen, displayed away from the mouse cursor, is closed when the mouse cursor moves over it."""
	foreground, background = _display_splash_screen(
		monkeypatch, [OUT_OF_SPLASH_SCREEN, OUT_OF_SPLASH_SCREEN, IN_SPLASH_SCREEN])
	frames = [_displayed_text_lines(frame) for frame in foreground.frames]
	assert [text_lines[23] for text_lines in frames] == ['', 'To continue: move the mouse cursor over this window']
	assert foreground.closed and background.closed


def test_splash_screen_closed_after_an_error(monkeypatch):
	"""
	Tests that the overlays of the splash screen are closed, even after an error.

	Before, they were only cleared: two more overlays stayed open at each display of the splash screen ("About").
	"""
	def get_cursor_pos():
		raise RuntimeError("GetCursorPos failed")
	overlays = _patch_splash_screen(monkeypatch, get_cursor_pos)
	with pytest.raises(RuntimeError):
		app.display_splash_screen()
	assert len(overlays) == 2 and all(overlay.closed for overlay in overlays)


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
	recorder. The recorder knows where the tray icon is, so that it does not record the clicks on it.
	"""
	happened = []
	trays = []
	fake_recorder = _FakeRecorderForMenu()
	fake_recorder.quit = lambda: fake_recorder.calls.append("quit")

	class FakeSysTrayIcon:
		def __init__(self, icon, hover_text, menu, on_quit=None, default_action=None):
			self.menu, self.on_quit, self.default_action = menu, on_quit, default_action
			trays.append(self)

		def contains(self, x, y):
			return False

		def run(self):
			happened.append("tray icon")

	monkeypatch.setattr(app, "Recorder", lambda: happened.append("recorder") or fake_recorder)
	monkeypatch.setattr(app, "display_splash_screen", lambda: happened.append("splash screen"))
	monkeypatch.setattr(app, "SysTrayIcon", FakeSysTrayIcon)
	assert app.run_interactive(app.parse_arguments(argv)) == app.EXIT_OK
	assert happened == events
	tray = trays[0]
	assert fake_recorder.is_on_tray_icon == tray.contains
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
	Tests that deleting the tooltip stops its process, or terminates it if it does not stop.

	The end of the process is waited for at most TOOLTIP_STOP_TIMEOUT seconds. Before, the recorder waited for it
	without a time limit.
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
	Tests that the process of the tooltip ends when it is stopped before its loop starts.

	quit() of its overlay is not called: quit() of overlay_arrows_and_more never returns if the window of the overlay
	is not created yet.
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


################################################################################################
#                   Recorded script and search of the element under the cursor                 #
################################################################################################

class _Rectangle(tuple):
	"""Replaces the rectangle of an element: (left, top, right, bottom)."""

	left = property(lambda self: self[0])
	top = property(lambda self: self[1])
	right = property(lambda self: self[2])
	bottom = property(lambda self: self[3])

	def mid_point(self):
		"""Returns the center of the rectangle."""
		return (self.left + self.right) // 2, (self.top + self.bottom) // 2

	def width(self):
		"""Returns the width of the rectangle."""
		return self.right - self.left

	def height(self):
		"""Returns the height of the rectangle."""
		return self.bottom - self.top


@pytest.mark.parametrize("dx, dy, relative_coordinate_mode, expected", [
	(0.5, -0.25, True, "%(50.00,-25.00)"),
	(0.5, -0.25, False, ""),
	(0.00004, 0.5, True, ""),  # rounded to 0.00
	(0.0, 0.5, True, ""),
])
def test_relative_coordinates(dx, dy, relative_coordinate_mode, expected):
	"""Tests the relative coordinates written after the path of an element in the recorded scripts."""
	assert recorder._relative_coordinates(dx, dy, relative_coordinate_mode) == expected


def test_get_script():
	"""Tests the script written for each type of event, and the UIPath blocks of the window and of the region."""
	events = [
		recorder.ClickEvent("left", 1, "Calculator||Window->Number pad||Group->One||Button", 0.5, -0.25, 1.0),
		recorder.ClickEvent("left", 2, "Calculator||Window->Number pad||Group->Two||Button", 0.0, 0.0, 2.0),
		recorder.ClickEvent("right", 1, "Calculator||Window->Number pad||Group->Three||Button", 0.0, 0.0, 3.0),
		recorder.SendKeysEvent('"123"'),
		recorder.MouseWheelEvent(-2),
		recorder.FindEvent("Calculator||Window->Display||Text", 0.1, 0.2, 4.0),
		recorder.DragAndDropEvent(
			"Untitled - Paint||Window->Canvas||Pane", -0.5, -0.5, "Untitled - Paint||Window->Canvas||Pane", 0.5, 0.5),
		recorder.MenuEvent("Untitled - Paint||Window", "File->Save as")]
	assert recorder._get_script(events, relative_coordinate_mode=True) == (
		"# encoding: " + sys.getdefaultencoding() + "\n\n"
		"from pywinauto_recorder.player import *\n\n\n"
		'with UIPath(u"Calculator||Window"):\n'
		'\twith UIPath(u"Number pad||Group"):\n'
		'\t\tclick(u"One||Button%(50.00,-25.00)")\n'
		'\t\tdouble_click(u"Two||Button")\n'
		'\t\tright_click(u"Three||Button")\n'
		'\t\tsend_keys("123")\n'
		'\t\tmouse_wheel(-2)\n'
		'\twrapper = find(u"Display||Text%(10.00,20.00)")\n\n'
		'with UIPath(u"Untitled - Paint||Window"):\n'
		'\tdrag_and_drop(u"Canvas||Pane%(-50.00,-50.00)", u"Canvas||Pane%(50.00,50.00)")\n'
		'\tmenu_click(u"File->Save as")\n')


def test_process_events_click():
	"""Tests that the press and the release of a mouse button become a ClickEvent, at the place of the release."""
	rectangle = _Rectangle((0, 0, 102, 42))
	events = [
		recorder.mouse.MoveEvent(10, 10, 0.0), recorder.ElementEvent(None, rectangle, "W||Window->A||Button"),
		recorder.mouse.ButtonEvent("down", "left", 1.0), recorder.mouse.ButtonEvent("up", "left", 1.1)]
	recorder._process_events(events)
	assert events[2:] == [recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1)]
	assert len(events) == 3


def test_process_events_double_click():
	"""Tests that a double click becomes a single ClickEvent, with 2 clicks."""
	rectangle = _Rectangle((0, 0, 102, 42))
	events = [
		recorder.mouse.MoveEvent(51, 21, 0.0), recorder.ElementEvent(None, rectangle, "W||Window->A||Button"),
		recorder.mouse.ButtonEvent("down", "left", 1.0), recorder.mouse.ButtonEvent("up", "left", 1.1),
		recorder.mouse.ButtonEvent("double", "left", 1.2), recorder.mouse.ButtonEvent("up", "left", 1.3)]
	recorder._process_events(events)
	assert events[2:] == [recorder.ClickEvent("left", 2, "W||Window->A||Button", 0.0, 0.0, 1.3)]


def test_process_events_drag_and_drop():
	"""Tests that a mouse move while the button is pressed makes a DragAndDropEvent, from an element to another one."""
	events = [
		recorder.mouse.MoveEvent(10, 10, 0.0),
		recorder.ElementEvent(None, _Rectangle((0, 0, 102, 42)), "W||Window->A||Button"),
		recorder.mouse.ButtonEvent("down", "left", 1.0),
		recorder.mouse.MoveEvent(150, 20, 1.2), recorder.mouse.MoveEvent(250, 21, 1.3),
		recorder.ElementEvent(None, _Rectangle((200, 0, 302, 42)), "W||Window->B||Button"),
		recorder.mouse.ButtonEvent("up", "left", 1.4)]
	recorder._process_events(events)
	assert events[2:] == [
		recorder.DragAndDropEvent("W||Window->A||Button", -0.82, -0.55, "W||Window->B||Button", -0.02, 0.0)]


def _move(x, y):
	"""Returns a mouse move event."""
	return recorder.mouse.MoveEvent(x, y, 0.0)


def _button(event_type, time_=1.0, button="left"):
	"""Returns a mouse button event: "down", "double" (a quick press, see the 'mouse' library) or "up"."""
	return recorder.mouse.ButtonEvent(event_type, button, time_)


def _element(path, left, top, right, bottom):
	"""Returns the event of an element found under the mouse cursor by the loop of the recorder."""
	return recorder.ElementEvent(None, _Rectangle((left, top, right, bottom)), path)


_A = _element("W||Window->A||Button", 0, 0, 102, 42)
_B = _element("W||Window->B||Button", 200, 0, 302, 42)
_X = _element("W||Window->X||Button", 250, 250, 350, 350)


def _clicks(events):
	"""Returns the click and drag and drop events, after the processing of the recorded events."""
	recorder._process_events(events)
	return [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.DragAndDropEvent))]


@pytest.fixture
def default_mouse_thresholds(monkeypatch):
	"""Uses the default mouse thresholds of Windows, so that the tests do not depend on the settings of Windows."""
	monkeypatch.setattr(recorder, "_mouse_thresholds", lambda: recorder.DEFAULT_MOUSE_THRESHOLDS)


@pytest.mark.parametrize("x, y", [(13, 8), (14, 6), (6, 14)])
def test_process_events_click_when_the_mouse_moves_a_little(default_mouse_thresholds, x, y):
	"""
	Tests that a press and a release are a click, at the position of the press, when the mouse moves less than the drag
	threshold of Windows (4 pixels). Before, the smallest move made a drag and drop (issue #75).
	"""
	events = [_move(10, 10), _A, _button("down"), _move(x, y), _button("up", 1.1)]
	assert _clicks(events) == [recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1)]


@pytest.mark.parametrize("x, y, dx2, dy2", [(15, 10, -0.72, -0.55), (10, 15, -0.82, -0.3)])
def test_process_events_drag_and_drop_beyond_the_drag_threshold(default_mouse_thresholds, x, y, dx2, dy2):
	"""Tests that a move farther than the drag threshold makes a drag and drop, here from an element to itself."""
	events = [_move(10, 10), _A, _button("down"), _move(x, y), _button("up")]
	assert _clicks(events) == [
		recorder.DragAndDropEvent("W||Window->A||Button", -0.82, -0.55, "W||Window->A||Button", dx2, dy2)]


def test_mouse_thresholds_of_windows(monkeypatch):
	"""Tests that the mouse thresholds are the ones of Windows, or the default ones when Windows does not give them."""
	metrics = {
		recorder.win32con.SM_CXDRAG: 10, recorder.win32con.SM_CYDRAG: 12,
		recorder.win32con.SM_CXDOUBLECLK: 8, recorder.win32con.SM_CYDOUBLECLK: 6}
	monkeypatch.setattr(recorder.win32api, "GetSystemMetrics", lambda index: metrics[index])
	assert recorder._mouse_thresholds() == recorder.MouseThresholds(drag=(10, 12), double_click=(4, 3))
	monkeypatch.setattr(recorder.win32api, "GetSystemMetrics", lambda index: 0)
	assert recorder._mouse_thresholds() == recorder.DEFAULT_MOUSE_THRESHOLDS == ((4, 4), (2, 2))


@pytest.mark.parametrize("events", [
	[_move(300, 300), _X, _move(10, 10), _button("down"), _A, _button("up", 1.1)],
	[_move(300, 300), _X, _move(10, 10), _button("down"), _button("up", 1.1), _A],
	[_move(300, 300), _X, _move(10, 10), _button("down"), _button("up", 1.1), _move(12, 11), _A],
], ids=["found during the press", "found after the release", "found after a small move"])
def test_process_events_click_on_the_element_found_after_the_press(default_mouse_thresholds, events):
	"""
	Tests a click just after a quick move: the recorder finds the element under the mouse cursor after the press.

	The recorder needs at least 0.1 s to find the element under the mouse cursor. Before, the click was recorded on the
	previous element, here X, even if it does not contain the position of the click.
	"""
	assert _clicks(events) == [recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1)]


def test_process_events_click_on_an_element_inside_the_element_found_before(default_mouse_thresholds):
	"""
	Tests a click on a button just after a quick move over the pane that contains it.

	The element found before the press contains the position of the click, but it is the pane: the button, found after
	the press at this position, is inside it. Before, the click was recorded on the pane, like click(u"||Custom") in the
	Calculator of the issue #75.
	"""
	events = [
		_move(300, 200), _element("W||Window->||Custom", 0, 0, 500, 400), _move(10, 10), _button("down"),
		_button("up", 1.1), _element("W||Window->||Custom->A||Button", 0, 0, 102, 42)]
	assert _clicks(events) == [recorder.ClickEvent("left", 1, "W||Window->||Custom->A||Button", -0.82, -0.55, 1.1)]


def test_process_events_click_ignores_a_new_element_after_the_click(default_mouse_thresholds):
	"""
	Tests that an element found after the click at the same position is ignored when it is not inside the clicked
	element: the click may have opened it.
	"""
	events = [_move(10, 10), _A, _button("down"), _button("up", 1.1), _element("Dialog||Window->OK||Button", 0, 0, 90, 30)]
	assert _clicks(events) == [recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1)]


def test_process_events_click_on_a_previous_element_that_contains_the_press(default_mouse_thresholds):
	"""
	Tests a click when the last element found before the press does not contain its position: the mouse cursor crossed
	a small button of a list item, then went back on the text of the list item. The click is on the list item.
	"""
	events = [
		_move(50, 20), _element("W||Window->L||List->Calculator||ListItem", 0, 0, 200, 40),
		_move(190, 20), _element("W||Window->L||List->Calculator||ListItem->Delete||Button", 180, 5, 198, 35),
		_move(60, 20), _button("down"), _button("up", 1.1)]
	click, = _clicks(events)
	assert click.path == "W||Window->L||List->Calculator||ListItem"
	assert (click.dx, click.dy) == (pytest.approx(-40 / 99), 0.0)


def test_process_events_click_when_no_element_contains_the_press(default_mouse_thresholds):
	"""
	Tests that the click is on the last element found before the press, as before, when no element found around the
	press contains its position: with the relative coordinates, the replay clicks at the same position.

	The element found after the mouse moved away from the click is not used.
	"""
	events = [
		_move(300, 300), _X, _move(10, 10), _button("down"), _button("up", 1.1), _move(100, 100),
		_element("W||Window->Y||Pane", 0, 0, 500, 500)]
	click, = _clicks(events)
	assert click.path == "W||Window->X||Button"
	assert (click.dx, click.dy) == (pytest.approx(-290 / 49), pytest.approx(-290 / 49))


def test_process_events_two_quick_clicks_on_two_elements(default_mouse_thresholds):
	"""
	Tests that 2 quick clicks on 2 elements are 2 clicks.

	The 'mouse' library gives a press soon after a release as a double click ('double'), even on another element.
	Before, the 2 clicks were recorded as a drag and drop from the first element to the second one.
	"""
	events = [
		_move(10, 10), _A, _button("down", 1.0), _button("up", 1.1), _move(250, 20), _B, _button("double", 1.3),
		_button("up", 1.4)]
	assert _clicks(events) == [
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.ClickEvent("left", 1, "W||Window->B||Button", -0.02, -0.05, 1.4)]


def test_process_events_double_click_after_a_small_move(default_mouse_thresholds):
	"""Tests that 2 quick clicks are a double click when the mouse moves less than the double click distance."""
	events = [
		_move(51, 21), _A, _button("down", 1.0), _button("up", 1.1), _move(53, 23), _button("double", 1.2),
		_button("up", 1.3)]
	assert _clicks(events) == [recorder.ClickEvent("left", 2, "W||Window->A||Button", 0.0, 0.0, 1.3)]


@pytest.mark.parametrize("first_click", [
	[_button("down", 1.0, "right"), _button("up", 1.1, "right")],
	[_button("down", 1.0), _move(80, 30), _button("up", 1.1), _move(51, 21)],
], ids=["with another button", "after a drag and drop"])
def test_process_events_quick_press_that_is_not_a_double_click(default_mouse_thresholds, first_click):
	"""Tests that a quick press is a click when the previous click is with another button, or a drag and drop."""
	events = [_move(51, 21), _A] + first_click + [_button("double", 1.2), _button("up", 1.3)]
	assert _clicks(events)[-1] == recorder.ClickEvent("left", 1, "W||Window->A||Button", 0.0, 0.0, 1.3)


def test_process_events_drag_and_drop_positions(default_mouse_thresholds):
	"""
	Tests a drag and drop that starts on an element found during the press, and its position at the release.

	Before, the drag and drop started on the previous element, X, and the position of the drop was the one where the
	last element was found, (200, 10), instead of the position of the release.
	"""
	events = [
		_move(300, 300), _X, _move(10, 10), _button("down"), _A, _move(200, 10),
		_element("W||Window->B||Pane", 150, 0, 302, 42), _move(250, 21), _button("up")]
	assert _clicks(events) == [
		recorder.DragAndDropEvent("W||Window->A||Button", -0.82, -0.55, "W||Window->B||Pane", 0.32, 0.0)]


def test_process_events_drop_ignores_the_elements_found_after_the_release(default_mouse_thresholds):
	"""Tests that the destination is not an element found after the drop: it may be the dragged element."""
	events = [
		_move(10, 10), _element("Desktop||Pane->D||List->A||ListItem", 0, 0, 102, 42), _button("down"), _move(250, 21),
		_element("Desktop||Pane->D||List", 0, 0, 500, 400), _button("up"),
		_element("Desktop||Pane->D||List->A||ListItem", 200, 0, 302, 42)]
	drag_and_drop, = _clicks(events)
	assert (drag_and_drop.path, drag_and_drop.path2) == ("Desktop||Pane->D||List->A||ListItem", "Desktop||Pane->D||List")


@pytest.mark.parametrize("events, nb_clicks", [
	([_move(10, 10), _A, _button("up")], 0),
	([_move(10, 10), _A, _button("down"), _button("up"), _button("up")], 1),
])
def test_process_events_ignores_a_release_without_press(default_mouse_thresholds, events, nb_clicks):
	"""
	Tests that a release without press is ignored: the button was pressed before the recording, or on the tray icon.

	Before, the processing of the recording raised an exception, or the release was merged with the previous click.
	"""
	assert len(_clicks(events)) == nb_clicks
	assert not [e for e in events if isinstance(e, recorder.mouse.ButtonEvent)]


def test_process_events_click_without_element(default_mouse_thresholds, capsys):
	"""
	Tests that a click is ignored, with a message, when no element was found under the mouse cursor.

	Before, the processing of the recording raised an exception, and the recording was lost.
	"""
	events = [_move(10, 10), _button("down"), _button("up"), _move(300, 300), _X, _button("down"), _button("up", 1.1)]
	assert _clicks(events) == [recorder.ClickEvent("left", 1, "W||Window->X||Button", 0.0, 0.0, 1.1)]
	assert capsys.readouterr().out == "A mouse click is not recorded: no element was found under the mouse cursor.\n"


_MODIFIER_SCAN_CODES = {"shift": 42, "ctrl": 29, "alt": 56, "alt gr": 541, "windows": 91}
"""The scan codes of the modifier keys in the keyboard events of Windows (541: the left ctrl sent with alt gr)."""


def _key(event_type, name, scan_code=0, is_keypad=False):
	"""Returns a keyboard event, like the 'keyboard' library gives it on Windows."""
	return recorder.keyboard.KeyboardEvent(event_type, scan_code, name=name, time=0.0, is_keypad=is_keypad)


def _press(name, scan_code=0, is_keypad=False):
	"""Returns the press and the release of a key."""
	return [_key("down", name, scan_code, is_keypad), _key("up", name, scan_code, is_keypad)]


def _typed(text):
	"""Returns the presses and the releases of the keys that type a text."""
	return [event for character in text for event in _press("space" if character == " " else character)]


def _held(modifier, events, name=None):
	"""Returns events while a modifier key is held. name: the name given by Windows in its language, like "maj"."""
	return [_key("down", name or modifier, _MODIFIER_SCAN_CODES[modifier])] + events + [
		_key("up", name or modifier, _MODIFIER_SCAN_CODES[modifier])]


_KEYBOARD_CASES = [
	pytest.param(_typed("2"), ['"2"'], id="key alone (issue #60)"),
	pytest.param(_typed("hello"), ['"hello"'], id="text without ENTER"),
	pytest.param(_held("shift", _typed("H")) + _typed("ello"), ['"Hello"'], id="shift"),
	pytest.param(_held("shift", _typed("H"), name="maj") + _typed("i"), ['"Hi"'], id="shift in French"),
	pytest.param(
		_typed("hello") + _press("tab", 15) + _typed("world") + _press("enter", 28), ['"hello{TAB}world{ENTER}"'],
		id="tab and enter"),
	pytest.param(_held("ctrl", _typed("s")), ['"{VK_CONTROL down}s{VK_CONTROL up}"'], id="ctrl+s"),
	pytest.param(_held("alt", _press("f4", 62)), ['"{VK_MENU down}{F4}{VK_MENU up}"'], id="alt+f4"),
	pytest.param(_held("shift", _press("tab", 15)), ['"{VK_SHIFT down}{TAB}{VK_SHIFT up}"'], id="shift+tab"),
	pytest.param(
		_held("ctrl", _held("shift", _press("esc", 1))),
		['"{VK_CONTROL down}{VK_SHIFT down}{ESC}{VK_SHIFT up}{VK_CONTROL up}"'], id="ctrl+shift+esc"),
	pytest.param(_typed("50%") + _press("enter", 28), ['"50{%}{ENTER}"'], id="percent"),
	pytest.param(
		_typed('{a}^b~c(d)e+f"g\\h'), ['"{{}a{}}{^}b{~}c(d)e+f\\"g\\\\h"'], id="special characters"),
	pytest.param(
		_press("suppr", 83) + _press("haut", 72) + _press("fin", 79) + _press("pg. suiv.", 81),
		['"{DELETE}{UP}{END}{PGDN}"'], id="keys named in French"),
	pytest.param(
		_typed("a") + _held("alt gr", _typed("@"), name="ctrl") + _typed("b.fr"), ['"a@b.fr"'], id="alt gr"),
	pytest.param(
		_typed("ab") + _held("ctrl", _typed("a")) + _held("shift", _typed("C")) + _typed("d"),
		['"ab{VK_CONTROL down}a{VK_CONTROL up}"', '"Cd"'], id="text after a shortcut"),
	pytest.param(_press("alt", 56) + _typed("f"), ['"{VK_MENU}f"'], id="alt alone"),
	pytest.param(_press("windows gauche", 91), ['"{LWIN}"'], id="windows alone"),
	pytest.param(
		_held("windows", _typed("e"), name="left windows"), ['"{VK_LWIN down}e{VK_LWIN up}"'], id="windows+e"),
	pytest.param(_press("ctrl", 29) + _press("shift", 42) + _typed("x"), ['"x"'], id="ctrl and shift alone"),
	pytest.param(
		_typed("a") + [_key("down", "ctrl", 29), _key("down", "alt", 56), _key("down", "s", 31), _key("up", "ctrl", 29),
		_key("up", "alt", 56)] + _typed("b"), ['"ab"'], id="shortcut of the recorder"),
	pytest.param(_typed("helo") + _press("backspace", 14) + _typed("lo"), ['"hello"'], id="backspace"),
	pytest.param(_press("backspace", 14) + _typed("ok"), ['"{BACKSPACE}ok"'], id="backspace of a previous text"),
	pytest.param(_press("caps lock", 58) + _typed("AB") + _press("verr.maj", 58), ['"AB"'], id="caps lock"),
	pytest.param(
		_press("1", 79, True) + _press("decimal", 83, True) + _press("end", 79, True), ['"1{VK_DECIMAL}{END}"'],
		id="numeric keypad"),
	pytest.param(
		_typed("a b") + _held("ctrl", _typed(" ")), ['"a b{VK_CONTROL down}{SPACE}{VK_CONTROL up}"'], id="space"),
	pytest.param(_typed("a") + _press("volume up", 48) + _typed("b"), ['"ab"'], id="unknown key"),
]
"""Keyboard events, and the strings of send_keys written in the script."""


@pytest.mark.parametrize("events, expected", _KEYBOARD_CASES)
def test_process_events_keyboard(events, expected):
	"""
	Tests the strings of send_keys written for keyboard events.

	Before, a text was lost when it was not followed by ENTER, like the keys that type no character, and the keys that
	type no character were written with names that send_keys does not know, like {f4} (issue #60).
	"""
	events = list(events)
	recorder._process_events(events)
	assert events == [recorder.SendKeysEvent(line) for line in expected]


def test_process_events_key_between_clicks(default_mouse_thresholds):
	"""Tests that a key pressed between 2 clicks is written, like a digit typed in the Calculator (issue #60)."""
	events = [_move(10, 10), _A, _button("down"), _button("up", 1.1)] + _typed("2") + [
		_move(210, 10), _B, _button("down", 2.0), _button("up", 2.1)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"2"'),
		recorder.ClickEvent("left", 1, "W||Window->B||Button", -0.82, -0.55, 2.1)]


def test_process_events_ctrl_click(default_mouse_thresholds):
	"""
	Tests that CTRL stays pressed during a click, and that a shortcut typed after the click presses CTRL again.

	send_keys sends a character as a virtual key only after the press or the release of a key in the same string.
	"""
	events = [_key("down", "ctrl", 29), _move(10, 10), _A, _button("down"), _button("up", 1.1)] + _typed("c") + [
		_key("up", "ctrl", 29)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.SendKeysEvent('"{VK_CONTROL down}"'),
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"{VK_CONTROL down}c{VK_CONTROL up}"')]


def test_process_events_releases_the_modifier_keys_at_the_end(default_mouse_thresholds):
	"""Tests that a modifier key still pressed at the end of the recording is released at the end of the script."""
	events = [_key("down", "ctrl", 29), _move(10, 10), _A, _button("down"), _button("up", 1.1)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.SendKeysEvent('"{VK_CONTROL down}"'),
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"{VK_CONTROL up}"')]


def test_process_events_keyboard_with_mouse_moves():
	"""Tests that the mouse moves between keyboard events do not cut the text in several strings of send_keys."""
	events = _typed("ab") + [_move(10, 10), _A] + _typed("c")
	recorder._process_events(events)
	assert events == [recorder.SendKeysEvent('"abc"'), _move(10, 10), _A]


@pytest.mark.skipif(sys.platform != "win32", reason="pywinauto needs Windows")
def test_send_keys_strings_are_parsed_by_pywinauto(monkeypatch):
	"""Tests that pywinauto parses the strings of send_keys written by the recorder, escaped by player.send_keys."""
	import ast
	import pywinauto.keyboard
	from pywinauto_recorder import player
	sent = []
	monkeypatch.setattr(pywinauto.keyboard, "send_keys", lambda keys, **kwargs: sent.append(keys))
	for case in _KEYBOARD_CASES:
		for line in case.values[1]:
			player.send_keys(ast.literal_eval(line))
	assert len(sent) == sum(len(case.values[1]) for case in _KEYBOARD_CASES)
	for keys in sent:
		assert pywinauto.keyboard.parse_keys(keys, with_spaces=True, with_tabs=True, with_newlines=True)


def test_compute_dx_dy_in_a_thin_rectangle():
	"""Tests that a relative coordinate is 0 in a rectangle that is too thin. Before, it was a division by zero."""
	assert recorder._compute_dx_dy(7, 1, _Rectangle((0, 0, 10, 2))) == (0.5, 0.0)
	assert recorder._compute_dx_dy(0, 7, _Rectangle((0, 0, 1, 10))) == (0.0, 0.5)


@pytest.mark.parametrize("path, common_path, relative_coordinate_mode, expected", [
	("A||Window->P||Pane->X||Button", "A||Window->P||Pane", True, "X||Button%(50.00,-25.00)"),
	("A||Window->P||Pane", "A||Window->P||Pane", False, ""),
	("A||Window->X||Button", "", False, "A||Window->X||Button"),
	("B||Window->Z||Button", "A||Window->P||Pane", True, "B||Window->Z||Button%(50.00,-25.00)~Absolute_UIPath"),
	("B||Window->Z||Button", "A||Window->P||Pane", False, "B||Window->Z||Button~Absolute_UIPath"),
	("A||Window->Q||Pane->Y||Button", "A||Window->P||Pane", False, "A||Window->Q||Pane->Y||Button~Absolute_UIPath"),
	("A||Window->P||Pane2->Y||Button", "A||Window->P||Pane", False, "A||Window->P||Pane2->Y||Button~Absolute_UIPath"),
	('B "1"||Window->Z||Button', "A||Window", False, 'B \\"1\\"||Window->Z||Button~Absolute_UIPath'),
])
def test_get_element_argument(path, common_path, relative_coordinate_mode, expected):
	"""
	Tests the path of an element written in a script: relative to the UIPath blocks, or else absolute.

	Before, the beginning of a path that is not under the path of the UIPath blocks was cut wrongly: the destination of
	a drag and drop in another window was "|Pane->||Pane->Type here to search||Button" in the issue #75.
	"""
	assert recorder._get_element_argument(path, 0.5, -0.25, common_path, relative_coordinate_mode) == expected


def test_get_script_of_a_drag_and_drop_to_another_window():
	"""Tests that the destination of a drag and drop in another window is written as an absolute path."""
	events = [
		recorder.DragAndDropEvent("A||Window->P||Pane->X||Button", 0.0, 0.0, "B||Window->Z||Button", 0.0, 0.0),
		recorder.ClickEvent("left", 1, "A||Window->P||Pane->Y||Button", 0.0, 0.0, 2.0)]
	assert recorder._get_script(events) == (
		"# encoding: " + sys.getdefaultencoding() + "\n\n"
		"from pywinauto_recorder.player import *\n\n\n"
		'with UIPath(u"A||Window"):\n'
		'\twith UIPath(u"P||Pane"):\n'
		'\t\tdrag_and_drop(u"X||Button", u"B||Window->Z||Button~Absolute_UIPath")\n'
		'\t\tclick(u"Y||Button")\n')


_ROW = "W||Window->T||Table->Row||Custom"
_CELL_WITH_REFERENCE = _ROW + "->Cell||Edit#[" + _ROW + "->Name||Text,2]"


@pytest.mark.parametrize("common_path, expected", [
	("W||Window->T||Table", "Row||Custom->Cell||Edit#[Row||Custom->Name||Text,2]"),
	(_ROW, "Cell||Edit#[Name||Text,2]"),  # before: "->Cell||Edit#[Name||Text,2]"
	("W||Window->H||Pane", "W||Window->T||Table->Row||Custom->Cell||Edit#["
		"W||Window->T||Table->Row||Custom->Name||Text~Absolute_UIPath,2]~Absolute_UIPath"),
])
def test_get_element_argument_with_a_reference_element(common_path, expected):
	"""
	Tests the path of an element found with a reference element (Smart mode): the path of the reference element is
	relative, or absolute, like the path of the element.
	"""
	assert recorder._get_element_argument(_CELL_WITH_REFERENCE, 0.0, 0.0, common_path, False) == expected


def test_get_element_argument_with_a_reference_element_in_another_branch():
	"""Tests that the path is absolute when the reference element is not under the path of the UIPath blocks."""
	path = "W||Window->T||Table->Cell||Edit#[W||Window->H||Pane->Name||Text,2]"
	assert recorder._get_element_argument(path, 0.0, 0.0, "W||Window->T||Table", False) == (
		"W||Window->T||Table->Cell||Edit#[W||Window->H||Pane->Name||Text~Absolute_UIPath,2]~Absolute_UIPath")


def _strategies(monkeypatch, smart_modes, unique_wrapper_path=None):
	"""Returns the strategies of a search of the same element, with the Smart mode given for each loop."""
	monkeypatch.setattr(recorder, "find_elements", lambda path, window_handle=None: ["element 1", "element 2"])
	search = recorder._ElementSearch()
	strategies = []
	for smart_mode in smart_modes:
		strategies.append(search.next_strategy("W||Window->B||Button", smart_mode).name)
		search.unique_wrapper_path = unique_wrapper_path
	return strategies


def test_element_search_strategies(monkeypatch):
	"""Tests the strategies tried while the mouse cursor stays on the same element and no unique path is found."""
	assert _strategies(monkeypatch, [False] * 4) == ["unique_path", "array_2d", "array_2d", "array_2d"]
	assert _strategies(monkeypatch, [True] * 4) == ["unique_path", "array_2d", "array_1d", "array_1d"]


def test_element_search_when_the_smart_mode_is_switched_off(monkeypatch):
	"""Tests that the search goes back to array_2d. Before, the loop of the recorder raised an IndexError."""
	assert _strategies(monkeypatch, [True, True, True, False, False]) == [
		"unique_path", "array_2d", "array_1d", "array_2d", "array_2d"]


def test_element_search_keeps_a_unique_path(monkeypatch):
	"""Tests that the strategy does not change when a unique path is found."""
	assert _strategies(monkeypatch, [True] * 3, unique_wrapper_path="W||Window->B||Button") == ["unique_path"] * 3


def test_element_search_starts_again_on_another_element(monkeypatch):
	"""Tests that the search starts again with the first strategy, and searches the elements, on another element."""
	searched_paths = []
	monkeypatch.setattr(
		recorder, "find_elements", lambda path, window_handle=None: searched_paths.append(path) or ["a", "b"])
	search = recorder._ElementSearch()
	assert search.next_strategy("W||Window->A||Button", False) == recorder.Strategy.unique_path
	assert search.next_strategy("W||Window->A||Button", False) == recorder.Strategy.array_2d
	assert search.next_strategy("W||Window->B||Button", False) == recorder.Strategy.unique_path
	assert searched_paths == ["W||Window->A||Button", "W||Window->B||Button"]
	assert search.elements == ["a", "b"]


def test_element_search_in_the_window_under_the_mouse_cursor(monkeypatch):
	"""
	Tests that the elements are searched in the window under the mouse cursor, and searched again in another window
	that has the same title.
	"""
	searches = []
	monkeypatch.setattr(
		recorder, "find_elements", lambda path, window_handle=None: searches.append((path, window_handle)) or ["a", "b"])
	search = recorder._ElementSearch()
	assert search.next_strategy("N||Window->B||Button", False, 1) == recorder.Strategy.unique_path
	assert search.next_strategy("N||Window->B||Button", False, 1) == recorder.Strategy.array_2d
	assert search.next_strategy("N||Window->B||Button", False, 2) == recorder.Strategy.unique_path
	assert searches == [("N||Window->B||Button", 1), ("N||Window->B||Button", 2)]


class _FakeOverlayForTheLoop:
	"""Replaces the overlay of the recorder: it keeps what is drawn."""

	def __init__(self):
		self.shapes = []

	def clear_all(self):
		self.shapes = []

	def add(self, **kwargs):
		self.shapes.append(kwargs)

	def refresh(self):
		pass


def _recorder_for_the_loop(wrapper, mode=recorder.RecorderMode.RECORD):
	"""
	Returns a Recorder whose thread and tooltip are not started, to test its loop.

	:param wrapper: The element under the mouse cursor, given by its desktop
	"""
	fake_recorder = object.__new__(recorder.Recorder)
	fake_recorder._mode = mode
	fake_recorder._smart_mode = False
	fake_recorder._copy_count = 0
	fake_recorder.main_overlay = _FakeOverlayForTheLoop()
	fake_recorder.desktop = types.SimpleNamespace(from_point=lambda x, y: wrapper)
	fake_recorder.event_list = [recorder.mouse.MoveEvent(0, 0, 0.0)]
	fake_recorder.last_element_event = None
	fake_recorder._window_handle = None
	fake_recorder._cursor_position_unavailable = False
	return fake_recorder


class _SizedElement:
	"""An element found by a search: it has a rectangle, and it is visible."""

	def __init__(self, rectangle):
		self._rectangle = _Rectangle(rectangle)

	def rectangle(self):
		return self._rectangle

	def is_visible(self):
		return True


class _VanishedElement:
	"""An element found by a search that does not exist anymore: UI Automation raises UIA_E_ELEMENTNOTAVAILABLE."""

	def rectangle(self):
		raise recorder.COMError(-2147220991, "An event was unable to invoke any of the subscribers", None)

	def is_visible(self):
		return self.rectangle()


class _FakeCachedFindElements:
	"""Replaces the find_elements of the recorder: it gives its results one after the other, and has a cache."""

	def __init__(self, *results):
		self.results = list(results)
		self.calls = []
		self.cache_clears = 0

	def __call__(self, path, window_handle=None):
		self.calls.append((path, window_handle))
		return self.results.pop(0) if len(self.results) > 1 else self.results[0]

	def cache_clear(self):
		self.cache_clears += 1


def test_loop_does_not_restrict_the_searches_of_the_player(monkeypatch):
	"""
	Tests that the loop of the recorder does not set the window where the player searches the elements.

	Before, it set the window under the mouse cursor: a replay from the tray menu, or from the command line, searched
	the elements only in this window, and failed in the other windows.
	"""
	from pywinauto_recorder import core
	monkeypatch.setattr(core, "native_window_handle", None)
	monkeypatch.setattr(recorder.win32api, "GetCursorPos", lambda: (1, 2))
	monkeypatch.setattr(recorder, "get_wrapper_path", lambda wrapper: "")  # no path: the loop ends at once
	monkeypatch.setattr(recorder.time, "sleep", lambda seconds: None)
	fake_recorder = _recorder_for_the_loop(types.SimpleNamespace(handle=5, parent=lambda: None))
	fake_recorder._Recorder__loop(recorder._ElementSearch(), 1)
	assert fake_recorder._window_handle == 5
	assert core.get_native_window_handle() is None


def test_loop_searches_again_when_the_elements_are_the_ones_of_a_previous_search(monkeypatch):
	"""
	Tests the loop of the recorder when the elements found are the ones of a previous search, kept by the cache: the
	application changed, and the element under the mouse cursor is not among them.

	The elements are searched again without the cache, and the element is found with the [] operator. Before, the path,
	which matches 2 elements, was recorded as a unique path, and the replay failed (issue #68).
	"""
	under_the_cursor = _Rectangle((60, 0, 110, 20))
	previous_page = [_SizedElement((0, 300, 50, 320)), _SizedElement((60, 300, 110, 320))]
	find_elements = _FakeCachedFindElements(previous_page, [_SizedElement((0, 0, 50, 20)), _SizedElement(under_the_cursor)])
	monkeypatch.setattr(recorder, "find_elements", find_elements)
	monkeypatch.setattr(recorder, "get_wrapper_path", lambda wrapper: "W||Window->B||Button")
	monkeypatch.setattr(recorder, "_is_cursor_in", lambda rectangle: True)
	monkeypatch.setattr(recorder.win32api, "GetCursorPos", lambda: (70, 10))
	wrapper = types.SimpleNamespace(handle=5, parent=lambda: None, rectangle=lambda: under_the_cursor)
	fake_recorder = _recorder_for_the_loop(wrapper)
	search = recorder._ElementSearch()
	fake_recorder._Recorder__loop(search, 1)  # Strategy.unique_path: the elements are searched again, 2 elements
	fake_recorder._Recorder__loop(search, 2)  # Strategy.array_2d
	assert find_elements.cache_clears == 1
	assert fake_recorder.event_list[1:] == [
		recorder.ElementEvent(recorder.Strategy.array_2d, under_the_cursor, "W||Window->B||Button#[0,1]")]


def test_element_search_again_without_the_cache(monkeypatch):
	"""Tests that the elements are searched again without the cache once for each element under the mouse cursor."""
	find_elements = _FakeCachedFindElements(["old 1", "old 2"], ["new"])
	monkeypatch.setattr(recorder, "find_elements", find_elements)
	search = recorder._ElementSearch()
	search.next_strategy("W||Window->B||Button", False, 7)
	search.i_strategy = 1
	assert search.search_again(_Rectangle((0, 0, 50, 20)))
	assert (search.elements, search.i_strategy, find_elements.cache_clears) == (["new"], 0, 1)
	assert find_elements.calls == [("W||Window->B||Button", 7)] * 2
	assert not search.search_again(_Rectangle((0, 0, 50, 20)))  # once for the element under the mouse cursor
	assert search.search_again(_Rectangle((60, 0, 110, 20)))  # another element that has the same path
	search.next_strategy("W||Window->C||Button", False, 7)
	assert search.search_again(_Rectangle((60, 0, 110, 20)))  # an element that has another path
	assert find_elements.cache_clears == 3


_UNDER_THE_CURSOR = (0, 0, 50, 20)
_ELSEWHERE = (0, 30, 50, 50)
_PATH = "W||Window->B||Button"


@pytest.mark.parametrize("rectangles, strategy, expected", [
	([_UNDER_THE_CURSOR], recorder.Strategy.unique_path, (recorder.Strategy.unique_path, _PATH, True)),
	([_ELSEWHERE], recorder.Strategy.unique_path, (recorder.Strategy.unique_path, _PATH, False)),
	([], recorder.Strategy.unique_path, (recorder.Strategy.unique_path, None, False)),
	([_UNDER_THE_CURSOR, (60, 0, 110, 20)], recorder.Strategy.unique_path, (recorder.Strategy.unique_path, None, True)),
	([_ELSEWHERE, (60, 30, 110, 50)], recorder.Strategy.unique_path, (recorder.Strategy.unique_path, None, False)),
	([_UNDER_THE_CURSOR, (60, 0, 110, 20)], recorder.Strategy.array_2d, (recorder.Strategy.array_2d, _PATH + "#[0,0]", True)),
	([_ELSEWHERE, (60, 30, 110, 50)], recorder.Strategy.array_2d, (recorder.Strategy.array_2d, None, False)),
	([(0, 0, 6, 6), (10, 0, 16, 6)], recorder.Strategy.array_2d, (recorder.Strategy.array_2d, None, False)),  # too small
	([], recorder.Strategy.array_2d, (recorder.Strategy.array_2d, None, False)),
])
def test_search_unique_path(monkeypatch, rectangles, strategy, expected):
	"""
	Tests the search of a unique path, and that it tells when the element under the mouse cursor is not among the
	elements found: they may be the ones of a previous search.

	Before, the search with the [] operator gave the path itself when the element under the mouse cursor was not among
	the elements sorted in rows: a path that matches several elements was taken for a unique path (issue #68).
	"""
	monkeypatch.setattr(recorder, "_is_cursor_in", lambda rectangle: True)
	fake_recorder = _recorder_for_the_loop(None)
	elements = [_SizedElement(rectangle) for rectangle in rectangles]
	if rectangles == [(0, 0, 6, 6), (10, 0, 16, 6)]:
		under_the_cursor = _Rectangle((0, 0, 6, 6))
	else:
		under_the_cursor = _Rectangle(_UNDER_THE_CURSOR)
	assert fake_recorder._Recorder__search_unique_path(elements, strategy, _PATH, under_the_cursor) == expected


@pytest.mark.parametrize("strategy", [recorder.Strategy.unique_path, recorder.Strategy.array_2d])
def test_search_unique_path_with_a_vanished_element(monkeypatch, strategy):
	"""Tests that an element found that vanished is reported: the elements found may be the ones of a previous search."""
	monkeypatch.setattr(recorder, "_is_cursor_in", lambda rectangle: True)
	fake_recorder = _recorder_for_the_loop(None)
	elements = [_SizedElement(_UNDER_THE_CURSOR), _VanishedElement()]
	assert fake_recorder._Recorder__search_unique_path(elements, strategy, _PATH, _Rectangle(_UNDER_THE_CURSOR)) == (
		strategy, None, False)


def test_search_unique_path_when_the_mouse_cursor_left_the_element(monkeypatch):
	"""Tests that nothing is searched when the mouse cursor left the element."""
	monkeypatch.setattr(recorder, "_is_cursor_in", lambda rectangle: False)
	fake_recorder = _recorder_for_the_loop(None)
	assert fake_recorder._Recorder__search_unique_path(
		[_VanishedElement()], recorder.Strategy.array_2d, _PATH, _Rectangle(_UNDER_THE_CURSOR)) == (None, None, True)


class _FakeAncestor:
	"""An element whose parent can be asked: each call is a request to the application."""

	calls = 0

	def __init__(self, parent=None):
		self._parent = parent

	def parent(self):
		_FakeAncestor.calls += 1
		return self._parent


@pytest.mark.parametrize("depth", [0, 1, 2, 5])
def test_top_level_window(depth):
	"""
	Tests the top level window of an element: its ancestor whose parent is the desktop.

	The parent of each ancestor is asked once. Before, 4 times, at each loop of the recorder (issue #68).
	"""
	desktop = _FakeAncestor()
	window = element = _FakeAncestor(desktop)
	for _ in range(depth):
		element = _FakeAncestor(element)
	_FakeAncestor.calls = 0
	assert recorder._top_level_window(element) is window
	assert _FakeAncestor.calls == depth + 2
	assert recorder._top_level_window(desktop) is desktop


class _FakeWin32Error(Exception):
	"""Replaces pywintypes.error, raised by the functions of win32api."""


def test_loop_when_the_position_of_the_mouse_cursor_is_not_available(monkeypatch, capsys):
	"""
	Tests the loop while Windows does not give the position of the mouse cursor: locked session, UAC prompt or
	CTRL+ALT+DEL screen. A message is printed once, and the loop waits. Before, an error was printed at each loop
	(issue #68).
	"""
	positions = [_FakeWin32Error(5, "GetCursorPos", "Access is denied."), _FakeWin32Error(5, "GetCursorPos", ""), (1, 2)]

	def get_cursor_pos():
		position = positions.pop(0)
		if isinstance(position, Exception):
			raise position
		return position

	monkeypatch.setattr(recorder, "win32api", types.SimpleNamespace(GetCursorPos=get_cursor_pos, error=_FakeWin32Error))
	monkeypatch.setattr(recorder, "get_wrapper_path", lambda wrapper: "")  # no path: the loop ends at once
	sleeps = []
	monkeypatch.setattr(recorder.time, "sleep", sleeps.append)
	fake_recorder = _recorder_for_the_loop(types.SimpleNamespace(handle=5, parent=lambda: None))
	for i in range(3):
		fake_recorder._Recorder__loop(recorder._ElementSearch(), i)
	assert capsys.readouterr().out == (
		"The position of the mouse cursor is not available (locked session, UAC prompt or CTRL+ALT+DEL screen): "
		"the recorder waits.\n")
	assert sleeps[:2] == [0.1, 0.1]
	assert (fake_recorder._cursor_position_unavailable, fake_recorder._window_handle) == (False, 5)


def test_is_cursor_in_when_the_position_of_the_mouse_cursor_is_not_available(monkeypatch):
	"""Tests that the mouse cursor is not in an element when Windows does not give its position."""
	def get_cursor_pos():
		raise _FakeWin32Error(5, "GetCursorPos", "Access is denied.")

	monkeypatch.setattr(recorder, "win32api", types.SimpleNamespace(GetCursorPos=get_cursor_pos, error=_FakeWin32Error))
	assert recorder._is_cursor_in(_Rectangle((0, 0, 100, 100))) is False


@pytest.mark.parametrize("error, vanished", [
	(recorder.COMError(-2147220991, "An event was unable to invoke any of the subscribers", (None, None, None, 0, None)), True),
	(recorder.COMError(0x80040201, "", None), True),
	(recorder.COMError(-2147467259, "Unspecified error", None), False),
	(RuntimeError("other error"), False),
])
def test_is_vanished_element_error(error, vanished):
	"""
	Tests the error of UI Automation for an element that vanished (UIA_E_ELEMENTNOTAVAILABLE): the loop of the recorder
	does not print it. Its message is misleading: "An event was unable to invoke any of the subscribers" (issue #68).
	"""
	assert recorder._is_vanished_element_error(error) is vanished


def _fake_recorder_for_mouse(mode=recorder.RecorderMode.RECORD, tray_icon=None):
	"""
	Returns an object that has the attributes used by Recorder.__mouse_on, without starting a recorder.

	:param tray_icon: The rectangle of the tray icon, (left, top, right, bottom), or None
	"""
	fake_recorder = types.SimpleNamespace(mode=mode, event_list=[], _mouse_position=None, _ignored_buttons=set())
	fake_recorder.is_on_tray_icon = recorder._no_tray_icon if tray_icon is None else (
		lambda x, y: tray_icon[0] <= x < tray_icon[2] and tray_icon[1] <= y < tray_icon[3])
	fake_recorder._Recorder__is_ignored = lambda event: recorder.Recorder._Recorder__is_ignored(fake_recorder, event)
	return fake_recorder


def _mouse_on(fake_recorder, events):
	"""Gives mouse events to Recorder.__mouse_on, and returns the recorded events."""
	for event in events:
		recorder.Recorder._Recorder__mouse_on(fake_recorder, event)
	return fake_recorder.event_list


def test_presses_on_the_tray_icon_are_not_recorded(monkeypatch):
	"""
	Tests that the presses on the tray icon, and their releases, are not recorded: they start or stop the recording.

	Before, the recorded script ended with a right click on the notification area, and its replay failed on it
	(issue #75).
	"""
	monkeypatch.setattr(recorder, "_is_menu_of_this_process", lambda x, y: False)
	fake_recorder = _fake_recorder_for_mouse(tray_icon=(100, 200, 124, 240))
	events = [
		_move(110, 210), _button("down", button="right"), _move(130, 210), _button("up", button="right"),
		_button("down"), _button("up")]
	assert _mouse_on(fake_recorder, events) == [_move(130, 210), _button("down"), _button("up")]


def test_presses_on_a_menu_of_the_recorder_are_not_recorded(monkeypatch):
	"""
	Tests that the presses on a menu of the process of the recorder, like the menu of the tray icon, and their releases,
	are not recorded. Before, a click on "Stop recording" was recorded.
	"""
	monkeypatch.setattr(recorder, "_is_menu_of_this_process", lambda x, y: (x, y) == (300, 400))
	fake_recorder = _fake_recorder_for_mouse()
	events = [_move(300, 400), _button("down"), _button("up"), _move(10, 10), _button("down"), _button("up")]
	assert _mouse_on(fake_recorder, events) == [_move(10, 10), _button("down"), _button("up")]


def test_mouse_events_are_recorded_only_in_record_mode(monkeypatch):
	"""Tests that the mouse events are not recorded outside the Record mode, and that the mouse position is kept."""
	monkeypatch.setattr(recorder, "_is_menu_of_this_process", lambda x, y: False)
	fake_recorder = _fake_recorder_for_mouse(mode=recorder.RecorderMode.INFO)
	assert _mouse_on(fake_recorder, [_move(5, 6), _button("down"), _button("up")]) == []
	assert fake_recorder._mouse_position == (5, 6)


@pytest.mark.parametrize("class_name, process_id, is_menu", [
	("#32768", "this process", True),
	("#32768", 1, False),
	("Button", "this process", False),
])
def test_is_menu_of_this_process(monkeypatch, class_name, process_id, is_menu):
	"""Tests the detection of the menus of the process of the recorder, like the menu of its tray icon."""
	import os
	if process_id == "this process":
		process_id = os.getpid()
	monkeypatch.setattr(recorder, "win32gui", types.SimpleNamespace(
		WindowFromPoint=lambda point: 7 if point == (300, 400) else 0,
		GetClassName=lambda window: class_name if window == 7 else ""))
	monkeypatch.setattr(recorder, "win32process", types.SimpleNamespace(
		GetWindowThreadProcessId=lambda window: (11, process_id)))
	assert recorder._is_menu_of_this_process(300, 400) is is_menu


def test_is_menu_of_this_process_when_windows_fails(monkeypatch):
	"""Tests that a position is not on a menu when Windows cannot give the window at this position."""
	def window_from_point(point):
		raise OSError("no window")

	monkeypatch.setattr(recorder, "win32gui", types.SimpleNamespace(WindowFromPoint=window_from_point))
	assert recorder._is_menu_of_this_process(300, 400) is False


def test_tray_icon_contains(monkeypatch):
	"""Tests that SysTrayIcon.contains asks Windows the rectangle of the icon, identified by its window and its id."""
	import ctypes
	identifiers = []
	results = [0, 0, 0, 0, 0x80004005]  # S_OK, then E_FAIL: Windows does not give the rectangle

	def shell_notify_icon_get_rect(identifier, rectangle):
		identifiers.append((identifier.contents.cbSize, identifier.contents.hWnd, identifier.contents.uID))
		rectangle.contents.left, rectangle.contents.top = 100, 200
		rectangle.contents.right, rectangle.contents.bottom = 124, 240
		return results.pop(0)

	monkeypatch.setattr(app, "windll", types.SimpleNamespace(shell32=types.SimpleNamespace(
		Shell_NotifyIconGetRect=shell_notify_icon_get_rect)))
	tray = object.__new__(app.SysTrayIcon)
	tray.hwnd = 42
	assert [tray.contains(100, 200), tray.contains(123, 239), tray.contains(124, 220), tray.contains(110, 199)] == [
		True, True, False, False]
	assert tray.contains(110, 210) is False
	assert identifiers[0] == (ctypes.sizeof(app._NOTIFYICONIDENTIFIER), 42, 0)


@pytest.mark.parametrize("text, char_width, expected", [
	("", 7.8, 25),
	("a\nb", 7.8, 25 + 16),
	("x" * 100, 6.8, 25 + 16),  # wrapped once: the tooltip is 500 pixels wide
	("x" * 100 + "\n" + "y" * 10, 6.8, 25 + 16 + 16),
])
def test_text_height_of_the_tooltip(text, char_width, expected):
	"""Tests the height of a box of the tooltip: a line more for each line break, and for each wrapped line."""
	from pywinauto_recorder import element_observer
	assert element_observer._text_height(text, char_width, 25) == expected


def test_tooltip_information_of_an_edit_box():
	"""Tests the text of an edit box in the tooltip, and that the tooltip works when it cannot be read."""
	from pywinauto_recorder import element_observer

	class FakeEdit:
		def __init__(self, text):
			self.text = text

		def text_block(self):
			if self.text is None:
				raise RuntimeError("UI Automation error")
			return self.text

	assert element_observer._add_type_info("info\n", FakeEdit("Hello"), "Name", "Edit") == (
		"info\n\nwrapper.text_block(): Hello\n")
	assert element_observer._add_type_info("info\n", FakeEdit("Name"), "Name", "Edit") == "info\n"
	assert element_observer._add_type_info("info\n", FakeEdit(None), "Name", "Edit") == "info\n"
