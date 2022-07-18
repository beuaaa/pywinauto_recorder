import pytest
import pywinauto
import time
import win32api
from pywinauto_recorder.core import Strategy


@pytest.fixture
def start_kill_app(request):
	"""
	`start_kill_app` starts an application and kills it when the test is done.
	
	:param request: The request with the application name for the application to start and to kill.
	"""
	app_name = request.param
	app = pywinauto.Application(backend="win32")
	app.start(app_name)
	yield app
	# 'app.connect' is needed by 'app.kill'
	# #time.sleep(2)  # Wait for the app to start because sometimes a process is spawned before the app is ready.
	# When the sleep duration is too short, 'app.connect' fails because of a spawned process with the same title.
	# Pywinauto 0.7.0 should be able to handle this without a sleep.
	app.connect(path=app_name, timeout=10, visible_only=True)
	app.kill(soft=True)


@pytest.fixture
def wait_recorder_ready():
	def recorder_ready(recorder, path_end, sleep=0.3):
		start_time = time.time()
		while l_e_e := recorder.get_last_element_event():
			x, y = win32api.GetCursorPos()
			if l_e_e.rectangle.top < y < l_e_e.rectangle.bottom:
				if l_e_e.rectangle.left < x < l_e_e.rectangle.right:
					if l_e_e.strategy == Strategy.unique_path:
						if path_end in l_e_e.path:
							break
			time.sleep(0.1)
		duration = time.time() - start_time
		if duration < sleep:
			time.sleep(sleep - duration)
	return recorder_ready
