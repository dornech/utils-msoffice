# prepare cloakbrowser for calling via VBA


"""
Download chromedriver and apply undetected-chromedriver patch to be called from other languages like VBA
return location via INI-file
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

from undetected_chromedriver.patcher import Patcher

import utils_mystuff as Utils



# TAP for parameters
class ParamsClass(TypedArgParse):
    params_inifile: str | None = None



# caller routines

# test routine for PythonIDE debugging - parameters are normally set via CLI
# hint: quotation marks surrounding parameter values are swallowed by CLI param transfer mechanism
def executeStandaloneTest() -> None:

    paramsparser = ParamsClass()
    params = paramsparser.parse_args()
    executeMain(params)


# VBA caller with direct CLI parameters (basically an CLI caller)
def executeVBAcallee() -> None:

    Utils.log_cli_args()
    paramsparser = ParamsClass(explicit_bool=True)
    params = paramsparser.parse_args()
    Utils.log_cli_params(params)

    executeMain(params)



#  main program logic

# main caller
def executeMain(params: ParamsClass) -> None:

    # determine ini file for parameters calling cloakbrowser externally
    if params.params_inifile is None:
        params.params_inifile = os.path.join(tempfile.gettempdir(), "undetectedchrome.ini")

    patcher = Patcher()
    patcher.auto()

    # determine content of ini file
    config = configparser.ConfigParser()
    config.add_section("undetectedchrome")
    config["undetectedchrome"]["binary_path"] = patcher.executable_path

    # write back ini file
    with open(params.params_inifile, 'w') as configfile:
        config.write(configfile)



# standalone call / test

if __name__ == '__main__':

    executeStandaloneTest()

    # completion message
    Utils.exitFinished("prepare_cloakbrowser")
