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



from importlib.metadata import PackageNotFoundError, version

# try:
#     __version__ = version('utils-msoffice')
# except PackageNotFoundError:  # pragma: no cover
#     __version__ = 'unknown'
# finally:
#     del version, PackageNotFoundError

# up-to-date version tag for modules installed in editable mode inspired by
# https://github.com/maresb/hatch-vcs-footgun-example/blob/main/hatch_vcs_footgun_example/__init__.py
# Define the variable '__version__':
try:

    # own developed alternative variant to hatch-vcs-footgun overcoming problem of ignored setuptools_scm settings
    # from hatch-based pyproject.toml libraries
    from hatch.cli import hatch
    from click.testing import CliRunner
    # determine version via hatch
    __version__ = CliRunner().invoke(hatch, ["version"]).output.strip()

except (ImportError, LookupError):
    # As a fallback, use the version that is hard-coded in the file.
    try:
        from ._version import __version__  # noqa: F401
    except ModuleNotFoundError:
        # The user is probably trying to run this without having installed the
        # package, so complain.
        raise RuntimeError(
            f"Package {__package__} is not correctly installed. Please install it with pip."
        )



import sys
import os.path

# switch os-path -> pathlib
sys.path.insert(1, os.path.dirname(os.path.realpath(__file__)))
# sys.path.insert(1, str(pathlib.Path(__file__).resolve().parent))

import utils_msoffice.utils_office as UtilsOffice
import utils_msoffice.utils_excel as UtilsExcel
import utils_msoffice.runner_VBA as RunnerVBA
