from setuptools import setup

import os.path
import sys
import codecs


def read(rel_path):
    here = os.path.abspath(os.path.dirname(__file__))
    with codecs.open(os.path.join(here, rel_path), 'r') as fp:
        return fp.read()


def get_version(rel_path):
    for line in read(rel_path).splitlines():
        if line.startswith('__version__'):
            separator = '"' if '"' in line else "'"
            return line.split(separator)[1]
    else:
        raise RuntimeError("Unable to find version string.")


# We need the path to setup.py to be able to run the setup from a different folder
def setup_path(path=""):
    """ Get the path to the setup file. """
    setup_path = os.path.abspath(os.path.split(__file__)[0])
    return os.path.join(setup_path, path)


sys.path.append(setup_path())   # add it to the system path

with open("README.rst", "r") as fh:
    long_description = fh.read()

setup(
    name='pywinauto_recorder',
    version=get_version("pywinauto_recorder/__init__.py"),
    packages=['pywinauto_recorder'],
    url='https://github.com/beuaaa/pywinauto_recorder',
    license='MIT',
    author='david pratmarty',
    author_email='david.pratmarty@gmail.com',
    description='Records/Replays GUI actions',
    long_description=long_description,
    long_description_content_type="text/x-rst",
    install_requires=['pywinauto', 'keyboard', 'mouse', 'overlay_arrows_and_more', 'enum34', 'pyperclip'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Console',
        'Environment :: Win32 (MS Windows)',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.8',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Software Development :: Testing',
        'Topic :: Software Development :: User Interfaces',
        'Topic :: Software Development :: Quality Assurance',
    ]
)
