# Utilities to work with MS Office applications from Python via COM interface technology

"""
Package with utilities to work with MS Office applications from Python via COM interface technology

Set of submodules contains:

- submodule for Microsoft Office constants -> deprecated, switched to access to win32.com.client.constants
- submodule for Microsoft VBA constants
- submodule with utilities for Microsoft Office in general
- submodule with utilities for Microsoft Excel including "pythonic" access to ExcelAPI
- submodule for interfacing calling COM scripting with Microsoft Office applications
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999

# fmt: off



# version determination

# original Hatchlor version
# from importlib.metadata import PackageNotFoundError, version
# try:
#     __version__ = version('{{ cookiecutter.project_slug }}')
# except PackageNotFoundError:  # pragma: no cover
#     __version__ = 'unknown'
# finally:
#     del version, PackageNotFoundError

# latest import requirement for hatch-vcs-footgun-example
from utils_msoffice.version import __version__


import sys
import os.path

# switch os-path -> pathlib
sys.path.insert(1, os.path.dirname(os.path.realpath(__file__)))
# sys.path.insert(1, str(pathlib.Path(__file__).resolve().parent))

import utils_msoffice.utils_office as UtilsOffice
import utils_msoffice.utils_excel as UtilsExcel
import utils_msoffice.runner_VBA as RunnerVBA
