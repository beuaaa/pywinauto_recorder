set PATH_PYTEST=C:\Users\oktalse\AppData\Local\Programs\python\Python38\Scripts\
REM %PATH_PYTEST%pytest.exe tests.py -m beuaaa_test --alluredir="./allure_result"
%PATH_PYTEST%pytest.exe tests.py --alluredir="./allure_result"

pause