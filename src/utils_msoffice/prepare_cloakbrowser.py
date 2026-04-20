# prepare cloakbrowser for calling via VBA

"""
Download cloakbrowser to be called from other languages like VBA
return location and call parameters via INI-file
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N802, N812
# others
# ruff: noqa: PLW1514

# fmt: off



import os
import tempfile
import configparser

from tap import Tap as TypedArgParse

from cloakbrowser.config import get_default_stealth_args
from cloakbrowser.download import ensure_binary

import utils_mystuff as Utils



# TAP for parameters
class ParamsClass(TypedArgParse):
    cloakbrowser_cache_dir: str | None = None
    params_inifile: str | None = None



# caller routines

# test routine for PythonIDE debugging - parameters are normally set via CLI
# hint: quotation marks surrounding parameter values are swallowed by CLI param transfer mechanism
def executeStandaloneTest() -> None:
    """
    execute standalone test
    """

    paramsparser = ParamsClass()
    params = paramsparser.parse_args()
    executeMain(params)


# VBA caller with direct CLI parameters (basically an CLI caller)
def executeVBAcallee() -> None:
    """
    callee from VBA (or other language via CLI)
    """

    Utils.log_cli_args()
    paramsparser = ParamsClass(explicit_bool=True)
    params = paramsparser.parse_args()
    Utils.log_cli_params(params)

    executeMain(params)



#  main program logic

# main caller
def executeMain(params: ParamsClass) -> None:
    """
    main routine to download/initialize CloakBrowser and return informationv ia INI-fiel to caller

    Args:
        params (): parameter dataclass filled via CLI
    """

    # determine cache path for cloakbrowser
    if params.cloakbrowser_cache_dir is None:
        params.cloakbrowser_cache_dir = os.path.join(tempfile.gettempdir(), "cloakbrowser")
    os.environ["CLOAKBROWSER_CACHE_DIR"] = params.cloakbrowser_cache_dir
    # determine ini file for parameters calling cloakbrowser externally
    if params.params_inifile is None:
        params.params_inifile = os.path.join(tempfile.gettempdir(), "cloakbrowser.ini")

    # determine content of ini file
    config = configparser.ConfigParser()
    config.add_section("cloakbrowser")
    config["cloakbrowser"]["binary_path"] = ensure_binary()
    config["cloakbrowser"]["stealth_args"] = ",".join(get_default_stealth_args())

    # write back ini file
    with open(params.params_inifile, 'w') as configfile:
        config.write(configfile)



# standalone call / test

if __name__ == '__main__':

    executeStandaloneTest()

    # completion message
    Utils.exitFinished("prepare_cloakbrowser")
