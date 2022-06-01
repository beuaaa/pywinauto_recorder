import pytest
import pywinauto
import time


@pytest.fixture
def run_app(request):
	app_name, window_name = request.param
	app = pywinauto.Application(backend="win32")
	app.start(app_name)
	time.sleep(1)
	app.connect(title=window_name, timeout=10)
	yield app
	app.kill(soft=True)
