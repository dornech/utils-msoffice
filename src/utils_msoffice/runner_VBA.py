# Utilities to work with MS Office applications from Python via COM interface technology
# core module for call interface for Python programs from CLI

"""
Module provides a call interface for Python programs from VBA (or any other language)
via command line call. For calling from COM applications a COM link can be established
to the calling application for bidirectional data exchange.
Parameter provisioning is supported via classic CLI calling or INI file.
Advantage is that complete parameter evaluation and logging is encapsuled in the Python
callee class.
"""


# ruff and mypy per file settings
#
# empty lines
# ruff: noqa: E302, E303
# naming conventions
# ruff: noqa: N801, N802, N803, N806, N812, N813, N815, N816, N818, N999
# boolean-type arguments
# ruff: noqa: FBT001, FBT002
# others
# ruff: noqa: E501, PLR0917, PLR1714, RUF013, SIM102
#
# disable mypy errors
# - mypy error "'object' has no attribute 'xyz' [attr-defined]" when accessing attributes of
#   dynamically bound wrapped COM object
# - mypy error "Returning Any from function ..."
# mypy: disable-error-code = "attr-defined, no-any-return"

# fmt: off



from typing import Callable, Optional, Union

import functools

import sys
import os
# from pathlib import Path

from tap import Tap as TypedArgParse

# switch os.path -> pathlib
sys.path.insert(1, os.path.dirname(os.path.realpath(__file__)))
# sys.path.insert(1, str(pathlib.Path(__file__).resolve().parent))

import utils_mystuff as Utils
import utils_msoffice.utils_office as UtilsOffice



class ErrorRunnerVBA(BaseException):
    pass



# TAP for minimum parameters -> to be synchronized with standardized VBA caller module

# ... base class with flag for call parameter logging only
class ParamsClassBase(TypedArgParse):
    """
    ParamsClassBase - argument parser base class for standardized CLI call interface
    """
    test_logcall_only: bool = False  # Flag for logging call only

# ... for call with com link
class ParamsClassCOMlinked(ParamsClassBase):
    """
    ParamsClassCOMlinked - extended ParamsClassBase with extension for COM link to calling Microsoft Office application
    """
    app: str  # Microsoft Office host application
    docfile: str  # Microsoft Office host document calling
    linkCOM: bool  # Flag for creating COM objects references for calling application (host + user application)

    def configure(self):
        choices_app = [
            "Excel",
            "Microsoft Excel",
            UtilsOffice.COMclass_Excel,
            "Access",
            "Microsoft Access",
            UtilsOffice.COMclass_Access
        ]
        self.add_argument("--app", choices=choices_app)

# ... for call via INI file
class ParamsClassINI(ParamsClassBase):
    """
    ParamsClassCOMlinked - extended ParamsClassBase with extension for parameter hand-over via INI file
    """
    inifile: str = ""  # INI file
    inisection: str = ""  # section in INI file to use as parameters

# ... for call via INI file and COM link
class ParamsClassCOMlinkedINI(ParamsClassINI, ParamsClassCOMlinked):
    """
    ParamsClassCOMlinkedINI - extended ParamsClassBase with extension for COM link and parameter hand-over via INI file
    """
    pass



# generalized caller object
class RunnerVBAcall:
    """
    RunnerVBAcall - runner object for calling Python from VBA (basically it is a
    caller server or Python callee intended for VBA originally but can be used
    otherwise as well)

    Calling Python from VBA is done via command line options. To provide an
    easy evaluation of the command line, the TypedArgumentParser package
    (TAP) is used. Basically the object avoids the need to copy some standard
    stuff  for parameter evaluation and logging into a Python program being called via
    a CLI interface. The main processing routine to be executed is "injected"
    into RunnerVBAcall.
    To use the object it is only necessary to define the TAP dataclass as the parameter
    interface and the main processing routine using this parameter interface.

    The CLI interface itself is implicitly defined via defining the parameter dataclass
    derived from the basic classes herein.

    For further comfort the object provides two standard call methods:
    - provisioning of all parameters via CLI (method executeVBAcallee)
    - provisioning of parameters for retrieval from INI file via CLI (method executeVBAcalleeINI)
    - provisioning parameters as list of parameters as argpase / TAP support (i.e.
      simulate CLI parameter handover)

    In addition, it allows a back-link to a COM host application - typically the
    calling application but not necessarily. The COM-linking allows to provide
    results back to the calling host application / document easily.

    To use the runner object, it is important the runner object is set up properly.
    To initialize, it is necessary to provide
    - the main execution routine containing the processing logic. All stuff around
      (i.e. parameter retrieval, COM-linking is done by the object itself).
      Signature must contain the params class parameter and if activated
      parameters for the calling COM host application, the COM document and
      a callback for controlling the statusbar of the COM host application.
    - flag to control/activate COM linking for calling the maine execution routine
    - parameter dataclass for normal CLI call
    - parameter dataclass for INI
    - callmethod for using __call__ interface of runtime object
    - flag to control logging

    The calling Python programm must use the object in one of the following
    ways (example assumes parameter provisioning via CLI completely,
    'executeMain_injected' as procedure/method containing main logic and
    'ParamClassXXX' to be the parameter dataclass derived from the respective
    RunnerVBA.ParamsClassXXX):

        # VBA caller with parameter retrieval from INI file provided as parameter
        def executeVBAcallerINI() -> None:

            # initialize object
            runner_object = RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = <True | False >, params_class = <ParamsClassCOMlinked | ParamsClass>, params_class_ini = <ParamsClassCOMlinkedINI | ParamsClassINI>)
            # call object method
            getattr(runner_object, RunnerVBA.RunnerVBAcall.executeVBAcalleeINI.__name__)()

        # VBA caller with direct CLI parameters (basically an CLI caller)
        def executeVBAcaller() -> None:

            # initialize object
            runner_object = RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = <True | False >, params_class = <ParamsClassCOMlinked | ParamsClass>)
            # call object method
            getattr(runner_object, RunnerVBA.RunnerVBAcall.executeVBAcallee.__name__)()
            # alternatively:
            RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = <True | False >, params_class = <ParamsClassCOMlinked | ParamsClass>)()

        # initialize object with call method for __call__ with parameter retrieval/logging
        executeRunnerVBA = RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >, params_class_ini = < ... >, callmethod = "executeVBAcalleeINI")
        executeRunnerVBA()
        executeRunnerVBA = RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >, callmethod = "executeVBAcallee")
        executeRunnerVBA()
        # alternatively:
        RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >, callmethod = "executeVBAcallee")()

        Argument parsing is supported as it is supported by argparse / TypedArgumentParser which might be helpful for
        testing i.e. following call works as well:
        RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >)([paramstr1, paramstr2, ... paramstr<n>n])

        # initialize object with call method for __call__ without parameter retrieval/logging but direct of injected executor with parameter class
        executeRunnerVBA = RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >, callmethod = "executeMain")
        executeRunnerVBA(params)
        # alternatively:
        RunnerVBA.RunnerVBAcall(execmain = executeMain_injected, linkCOMargs = < ... >, params_class = < ... >, callmethod = "executeMain")(params)

        # main caller - basic version
        def executeMain_injected(params: ParamsClass) -> None:

            # do stuff
            pass

        # main caller - COM-linked version
        def executeMain_injected(params: ParamsClass, app: object, doc: object, statuscallback: Callable) -> None:

            # do stuff
            pass
    """

    def __init__(
        self,
        execmain: Callable,
        linkCOMargs: bool = False,
        params_class: type[ParamsClassBase] = None,   # type: ignore
        params_class_ini: Optional[type[ParamsClassINI]] = None,
        callmethod: str = "",
        log: bool = True
    ):
        """
        __init__ - initialize VBA runner object

        Args:
            execmain (Callable): main routine, parameter must be defined as paramsclass and potentially COM Link args.
            linkCOMargs (bool, optional): True if signature of 'execmain' callable contains linked COM objects (application, document and statusbar callback). Defaults to False.
            params_class (Type[ParamsClassBase]): params dataclass for normal call and target structure for reading from INI.
            params_class_ini (Optional[Type[ParamsClassINI]], optional): params dataclass for parameter retrieval from INI file. Defaults to None.
            callmethod (str): method for direct call, must be valid object method. Defaults to empty string (but then set internally to "executeVBAcallee").
            log (bool, optional): Flag to control parameter logging. Defaults to True.
        """

        if params_class is None and params_class_ini is None:
            err_msg = "Parameter dataclass for retrieving parameters not set."
            raise AttributeError(err_msg)

        self._execmain = execmain
        self._linkCOMargs = linkCOMargs
        self._params_class = params_class
        self._params_class_ini = params_class_ini
        self._callmethod = None
        if callmethod != "":
            if hasattr(self, callmethod):
                self._callmethod = callmethod
        else:
            self._callmethod = self.executeVBAcallee.__name__
        self._log = log

    # __call__ is used to allow object to be called from outside if initialized by Python programm
    def __call__(self, args, *kwargs):

        if self._callmethod is not None:
            if hasattr(self, self._callmethod):
                getattr(self, self._callmethod)(args, *kwargs)
            else:
                raise AttributeError()

    @staticmethod
    def assignCOMobjects(params: ParamsClassCOMlinked) -> tuple[object, object, bool]:
        """
        assignCOMobjects - assign COM objects for COM link to calling host

        creating COM objects references for calling application (host + user application)
        # Assumption: office host application is already started (due to problem with ACCESS)

        Args:
            params (ParamsClassCOMlinked): argument parser object

        Returns:
            Tuple[object, object, bool]: application COM object, document COM object, flag application started by function
        """

        appCOMobj: object = None
        docCOMobj: object = None
        started_app = False

        if params.linkCOM is True:

            if (params.docfile is not None) and (params.docfile != "") and os.path.exists(params.docfile):

                try:
                    docCOMobj = UtilsOffice.assignCOMdocument(params.docfile)
                except BaseException:
                    # pass
                    err_msg = "Invalid docfile parameter."
                    raise ErrorRunnerVBA(err_msg)  # noqa: B904

                if docCOMobj is not None:
                    appCOMobj = docCOMobj.Parent

            if appCOMobj is None:

                if "Access".upper() in params.app.upper():
                    appCOMclass = UtilsOffice.COMclass_Access
                elif "Excel".upper() in params.app.upper():
                    appCOMclass = UtilsOffice.COMclass_Excel
                else:
                    err_msg = "Invalid identifier for Office Application used."
                    raise ErrorRunnerVBA(err_msg)

                appCOMobj, started_app = UtilsOffice.assignCOMapplication(appCOMclass, False)

                if (appCOMobj is not None) and (params.docfile != ""):
                    if appCOMclass == UtilsOffice.COMclass_Access:
                        if (
                            appCOMobj.CurrentProject.FullName == params.docfile or
                            appCOMobj.CurrentProject.Name == params.docfile
                        ):
                            docCOMobj = appCOMobj.CurrentProject
                        else:
                            err_msg = "Requested Office document not open in application."
                            raise ErrorRunnerVBA(err_msg)
                    elif appCOMclass == UtilsOffice.COMclass_Excel:
                        try:
                            docCOMobj = appCOMobj.Workbooks(params.docfile)
                        except BaseException:
                            err_msg = "Requested Office document not open in application."
                            raise ErrorRunnerVBA(err_msg)  # noqa: B904

        return appCOMobj, docCOMobj, started_app

    # read params from config file
    @staticmethod
    def readini2params(
        params: Union[ParamsClassINI, ParamsClassCOMlinkedINI], ParamsClass: type[ParamsClassBase]
    ) -> ParamsClassBase:
        """
        readini2params - read params from INI file into argument parser object

        Args:
            params (Union[ParamsClassINI, ParamsClassCOMlinkedINI]): calling params with INI file arguments
            ParamsClass (type[ParamsClassBase]): params class (not object instance, instantiation within function)

        Returns:
            ParamsClassBase: argument parser object

        target parameter object must be passed as class parameter!
        """

        # read ini-file (existence is checked in reader)
        inifile_config = Utils.readconfigfile(params.inifile, lambda option: option)
        # check section parameter
        if not inifile_config.has_section(params.inisection):
            err_msg = "INI section not provided or not valid."
            raise Exception(err_msg)
        inifile_configdict = {**inifile_config[params.inisection]}

        # check parameters - delete superfluous keys from ini-file
        # watch out: keys/entries cannot be deleted while looping over the dictionary, therefore loop
        # over keylist as temporary list
        for key in list(inifile_configdict.keys()):
            if key not in ParamsClass.__annotations__:
                del inifile_configdict[key]

        # check parameters - add/overwrite parameters by values provided via CLI
        params_dict = params.as_dict()
        # for key, value in params_dict:
        #     if key in ParamsClass.__annotations__:
        #         inifile_configdict[key] = value
        inifile_configdict = {key: params_dict.get(key, inifile_configdict[key]) for key in inifile_configdict}
        # Utils.copydictfields(params_dict, inifile_configdict)

        # arg-parse from ini-file params dictionary into params structure used for calling
        # note: mandatory parameters (i.e. app + file + linkCOM) must be contained in source dictionary
        paramsparser_from_INI = ParamsClass(explicit_bool=True)
        # params_from_INI = paramsparser_from_INI.from_dict(inifile_configdict)
        if issubclass(type(params), ParamsClassCOMlinkedINI):
            params_from_INI = paramsparser_from_INI.from_dict(
                {
                    "app": params.app,
                    "docfile": params.docfile,
                    "linkCOM": params.linkCOM, **inifile_configdict
                }
            )
        else:
            params_from_INI = paramsparser_from_INI.from_dict({**inifile_configdict})
        Utils.copydictfields(params_dict, params_from_INI)

        return params_from_INI

    def executeVBAcalleeINI(self, params_list: Optional[list[str]] = None) -> None:
        """
        VBA callee interface for parameter retrieval from INI file provided as parameter
        """

        if self._params_class_ini is not None and self._params_class is not None:

            # parse params
            paramsparser = self._params_class_ini(explicit_bool=True)
            if params_list is None:
                # read params from command line
                if self._log:
                    Utils.log_cli_args()
                params = paramsparser.parse_args()
            else:
                # read params from list of str
                params = paramsparser.parse_args(params_list)
            if self._log:
                Utils.log_cli_params(params)

            # read params from INI
            params_from_INI = self.readini2params(params, self._params_class)
            if self._log:
                Utils.log_cli_params(params_from_INI)

            if not params.test_logcall_only:
                self.executeMain(params_from_INI)

        else:

            err_msg = "Parameter dataclass for retrieving parameters from INI file not set."
            raise AttributeError(err_msg)

    def execute_VBAcallee_from_INI(self, params_list: Optional[list[str]] = None) -> None:
        self.executeVBAcalleeINI(params_list)

    # def executeVBAcallerINI(self, params_list: Optional[list[str]] = None) -> None:
    #     self.executeVBAcalleeINI(params_list)

    # def execute_VBAcaller_from_INI(self, params_list: Optional[list[str]] = None) -> None:
    #     self.executeVBAcalleeINI(params_list)

    def executeVBAcallee(self, params_list: Optional[list[str]] = None) -> None:
        """
        VBA callee interface with direct CLI parameters (basically an CLI callee)
        """

        if self._params_class is not None:

            # parse params
            paramsparser = self._params_class(explicit_bool=True)
            if params_list is None:
                # read params from command line
                if self._log:
                    Utils.log_cli_args()
                params = paramsparser.parse_args()
            else:
                # read params from list of str
                params = paramsparser.parse_args(params_list)
            if self._log:
                Utils.log_cli_params(params)

            if not params.test_logcall_only:
                self.executeMain(params)

        else:

            err_msg = "Parameter dataclass for retrieving parameters not set."
            raise AttributeError(err_msg)

    def execute_VBAcallee(self, params_list: Optional[list[str]] = None) -> None:
        self.executeVBAcallee(params_list)

    # def executeVBAcaller(self, params_list: Optional[list[str]] = None) -> None:
    #     self.executeVBAcallee(params_list)

    # def execute_VBAcaller(self, params_list: Optional[list[str]] = None) -> None:
    #     self.executeVBAcallee(params_list)

    def executeMain(self, params: Union[ParamsClassBase, ParamsClassCOMlinked]) -> None:
        """
        main routine executed, main processing is "injected" here
        """

        if not isinstance(params, self._params_class):
            err_msg = "Param object class does not match."
            raise ValueError(err_msg)

        if not params.test_logcall_only:

            if issubclass(type(params), ParamsClassCOMlinked):
                if params.linkCOM:
                    # initialize COM references / link COM object(s)
                    app: object = None
                    doc: object = None
                    started_app: bool = False
                    app, doc, started_app = self.assignCOMobjects(params)
                    # initialize status callback
                    statuscallback = functools.partial(UtilsOffice.set_app_status, appCOMobj=app)
                    # print(f"COM-Link aufgebaut: Anwendung {app.Name}, Datei {doc.Name}, started_app {started_app}")
                else:
                    # print(f"COM-Link nicht aktiviert.")
                    pass
            else:
                # print(f"COM-Link nicht parametrisiert.")
                pass

            if not self._linkCOMargs:
                self._execmain(params)
            else:
                self._execmain(params, app, doc, statuscallback)

            if issubclass(type(params), ParamsClassCOMlinked):
                if params.linkCOM:
                    # reset status
                    app.StatusBar = False
                    # close app
                    if started_app:
                        UtilsOffice.quit_started_app(app)

        # else:
        #
        #   print(params, "\n\n")

    def execute_main(self, params: ParamsClassBase) -> None:
        self.executeMain(params)

    def execMain(self, params: ParamsClassBase) -> None:
        self.executeMain(params)

    def exec_main(self, params: ParamsClassBase) -> None:
        self.executeMain(params)
