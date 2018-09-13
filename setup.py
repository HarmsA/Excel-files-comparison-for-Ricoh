from cx_Freeze import setup, Executable
import os.path
import sys

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

executables = [Executable("main.py")]


base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(  name='Ricoh TAS REPORT Program',
        version='0.1',
        description='Creates TAS Training Requests',
        author='Adam Harms',
        options={"build_exe": {"packages": ['numpy']}},
        executables = [Executable('main.py', base = base)])