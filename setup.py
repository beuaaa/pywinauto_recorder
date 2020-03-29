from setuptools import setup

import pywinauto_recorder

setup(
    name='pywinauto_recorder',
    version=pywinauto_recorder.__version__,
    packages=['pywinauto_recorder'],
    url='https://github.com/beuaaa/pywinauto_recorder',
    license='MIT',
    author='david pratmarty',
    author_email='david.pratmarty@gmail.com',
    description='pywinauto recorder', install_requires=['pywinauto', 'keyboard', 'mouse', 'overlay_arrows_and_more']
)
