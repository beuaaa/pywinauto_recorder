from setuptools import setup

with open("README.rst", "r") as fh:
    long_description = fh.read()

setup(
    name='pywinauto_recorder',
    version='0.1.9',
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
