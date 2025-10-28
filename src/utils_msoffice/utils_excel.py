# Utilities to work with MS Office applications from Python via COM interface technology
# Excel utilities

"""
Module provides several utilities for Excel and wrapper classes to enable more 'pythonic'
access to Excel API.


Example / doctest demonstrating relevant features:
```
>>> import os
>>> import tempfile
>>> import utils_msoffice as UtilsMSOffice

>>> # determine test files
>>> xlsfile1 = "PyFile1.xls"
>>> xlspath1 = os.path.join(tempfile.gettempdir(), xlsfile1)
>>> if os.path.exists(xlspath1): os.remove(xlspath1)
>>> xlsfile2 = "PyFile2.xls"
>>> xlspath2 = os.path.join(tempfile.gettempdir(), xlsfile2)
>>> if os.path.exists(xlspath2): os.remove(xlspath2)

>>> # start Excel wrapper
>>> xlapp = UtilsMSOffice.UtilsExcel.xlAppWrapper()
>>> # determine number of open workbooks
>>> wb_opened = xlapp.Workbooks.Count

>>> # open and save test workbooks
>>> wb1 = xlapp.Workbooks.Add()
>>> wb1.SaveAs(xlspath1)
>>> print("check saveas test workbook 1:", xlapp.ActiveWorkbook.Name == wb1.name == xlsfile1)
check saveas test workbook 1: True
>>> wb2 = xlapp.workbooks.add()
>>> wb2.save_as(file_name=xlspath2)
>>> print("check saveas test workbook 2:", xlapp.ActiveWorkbook.FullName == wb2.full_name == xlspath2)
check saveas test workbook 2: True

>>> # check getitem for worksheets
>>> print("check getitem for worksheets:", wb1[1].name == wb1.Worksheets(1).Name)
check getitem for worksheets: True

>>> # check __getitem__/__setitem__ for worksheets and parameters
>>> wb1[1][1,1] = 5
>>> wb1[1]["A2"] = 10
>>> wb1.close(SaveChanges=False)
>>> wb1 = xlapp.workbooks.open(xlspath1)
>>> print("check save of test workbook with savechanges=False:", wb1[1][1,1].value!=5)
check save of test workbook with savechanges=False: True
>>> wb1[1][1,1] = 5
>>> wb1[1]["A2"] = 10
>>> wb1.close(savechanges=True)
>>> wb1 = xlapp.workbooks.open(xlspath1)
>>> print("check save of test workbook with savechanges=True:", wb1[1][1,1].value==5)
check save of test workbook with savechanges=True: True

>>> # access un-wrapped vs wrapped EXCEL and item-access
>>> print("check wrapped/unwrapped access and item access:", xlapp._xlWrapped.Workbooks(wb_opened+1).Name == xlapp[wb_opened+1].name)
check wrapped/unwrapped access and item access: True

>>> # access via object hierarchy and pythonized identifiers
>>> xlCollObj = xlapp.Workbooks(wb_opened+2).Worksheets(1).Parent.Parent.Workbooks
>>> print(xlCollObj.__class__.__name__, xlCollObj.count == xlapp.Workbooks(wb_opened+1).Worksheets(1).parent.parent.workbooks.count == xlapp.work_books.count)
xlWorkbooksWrapper True

>>> # check de-wrapping of wrapped objects with add worksheet
>>> wb1_sheets = wb1.Worksheets.count
>>> ws_added = wb1.Worksheets.Add(After=wb1.worksheets[2], Count=2)
>>> print("added sheet:", xlapp[xlsfile1].worksheets.count == wb1_sheets+2)
added sheet: True

```
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
# ruff: noqa: A002, ICN001, E115, E402, E501, F841, PLR5501, RUF059, SIM102, UP007
#
# disable mypy errors
# - mypy error "'object' has no attribute 'xyz' [attr-defined]" when accessing attributes of
#   dynamically bound wrapped COM object
# - mypy error "Returning Any from function ..."
# mypy: disable-error-code = "attr-defined, no-any-return"

# fmt: off



from typing import Any, ClassVar, NewType, Optional, Union
from dataclasses import dataclass

import os.path
# import pathlib
import tempfile
import atexit
import datetime

import pythoncom

import functools
import multimethod

import numpy
import pandas

# import PySimpleGUI as simplegui
import FreeSimpleGUI as simplegui
simplegui.theme("Default1")

import utils_mystuff as Utils
import utils_msoffice.utils_office as UtilsOffice



# exception class
class ErrorUtilsExcel(BaseException):
    """Replaces pywintypes.com_error with more informative error messages."""
    pass



# class for storing Excel flags
@dataclass
class ExcelFlagsClass:
    """
    ExcelFlagsClass - class for storing Excel application flags
    """
    Calculation: int = 0
    EnableEvents: bool = False
    ScreenUpdating: bool = False
    initialized: bool = False



class xlGenericWrapper:
    """
    xlGenericWrapper - generic wrapper as pass through for calls to wrapped object

    The generic wrapper allows a more 'pythonic' access to the Microsoft Office
    application API without duplicating the API. Precondition is that the underlying
    application objects are wrapped by wrapper classes.

    To allow recursive resolution for application specific extended classes an own
    generic wrapper class is required per application. For Excel the wrapper
    classes 'xlAppWrapper', 'xlWorkbookWrapper', 'xlWorksheetWrapper' and
    'xlRangeWrapper' are defined and considered in the __gettattr__
    implementation.

    Example: the Excel method Workbook.SaveAs can be called from a
    'xlWorkbookWrapperClass' object as
    - xlWorkbook.SaveAs (standard method call)
    - xl_workbook_object_wrapped.SaveAs (standard method call)
    - xl_workbook_object_wrapped.saveas (method call but in lower case)
    - xl_workbook_object_wrapped.save_as ('pythonic' method in 'snake case')

    The mechanism also works recursively if wrappers are assigned along the path i. e.
    - xlRangeObject.Parent.Parent.Name
    - xl_range_object_wrapped.Parent.Parent.Name
    both return the name of the workbook.

    To overcome the problem to implement specific wrapper object for COM
    collection objects f. e. Workbooks(<idx>) methods in the application
    specific wrapper class are used to overload the COM object collection
    objects.

    The mechanism is applied to wrapper class methods extending the Excel-API as well.
    """

    _cls_attrmap: ClassVar[dict] = {}
    _cls_attrmap_wrapped_get: ClassVar[dict] = {}
    _cls_attrmap_wrapped_put: ClassVar[dict] = {}
    _cls_attrmap_wrapped_method: ClassVar[dict] = {}

    def __init__(self, xlWrapped):  # docsig: disable=SIG102

        # initialize instance - version with attribute mappings as class
        # attributes are initialized on first call only
        # - use of class method __new__ not possible due dependency on __init__ parameter
        # - mappings must be on class level because otherwise recursion error
        # - save time to fill attribute mappings via get_attrmap / get_attrmapCOM

        # initialize instance - mapping tables
        if self.__class__._cls_attrmap == {}:
            self.__class__._cls_attrmap = UtilsOffice.get_attrmap(self.__class__)
            # manually add relevant attribute
            self.__class__._cls_attrmap["_xlWrapped"] = "_xlWrapped"
            self.__class__._cls_attrmap["_xlwrapped"] = "_xlWrapped"
            self.__class__._cls_attrmap["xlwrapped"] = "_xlWrapped"
            self.__class__._cls_attrmap_wrapped_get, self.__class__._cls_attrmap_wrapped_put, self.__class__._cls_attrmap_wrapped_method = UtilsOffice.get_attrmapCOM(xlWrapped)

        # initialize instance - wrapped object
        self._xlWrapped = xlWrapped

    # pass through for Microsoft Office object model get-attributes to wrapped COM object (properties and methods)
    def __getattr__(self, attr) -> Any:

        # getattr for "pythonized" caller names
        # a) wrapper class methods
        # b) wrapped object methods and wrapped object methods for get
        # NOTE: directly available attributes/methods in wrapper class are caught via __getattribute__
        attr_lower = attr.lower()
        attr_desnaked = attr_lower.replace("_", "")
        if attr_lower in self.__class__._cls_attrmap:
            retval = getattr(self, self.__class__._cls_attrmap[attr_lower])
        elif attr_desnaked in self.__class__._cls_attrmap:
            retval = getattr(self, self.__class__._cls_attrmap[attr_desnaked])
        elif attr.lower() in self.__class__._cls_attrmap_wrapped_get:
            retval = getattr(self._xlWrapped, self.__class__._cls_attrmap_wrapped_get[attr_lower])
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_get:
            retval = getattr(self._xlWrapped, self.__class__._cls_attrmap_wrapped_get[attr_desnaked])
        elif attr.lower() in self.__class__._cls_attrmap_wrapped_method:
            retval = functools.partial(UtilsOffice.callwrapper_COMmethod, self._xlWrapped, self.__class__._cls_attrmap_wrapped_method[attr_lower], self._wrap_retval)
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_method:
            retval = functools.partial(UtilsOffice.callwrapper_COMmethod, self._xlWrapped, self.__class__._cls_attrmap_wrapped_method[attr_desnaked], self._wrap_retval)
        elif hasattr(self._xlWrapped, attr):
            if attr not in self.__class__._cls_attrmap_wrapped_method.values():
                retval = getattr(self._xlWrapped, attr)
            else:
                retval = functools.partial(UtilsOffice.callwrapper_COMmethod, self._xlWrapped, attr, self._wrap_retval)
        else:
            err_msg = f"'{self!r}' object has no attribute '{attr}'"
            raise AttributeError(err_msg)

        return self._wrap_retval(retval)

    # internal return object wrapper - must be preceeded by _ !!!
    @staticmethod
    def _wrap_retval(retval: object) -> object:

        # wrap return value for recursion if wrapper exists -> may cause stack overflow error in debugging
        typename = type(retval).__name__.replace("_", "")
        if typename == "Application":
            return xlAppWrapper(retval)
        elif typename == "Workbooks":
            return xlWorkbooksWrapper(retval)
        elif typename == "Workbook":
            return xlWorkbookWrapper(retval)
        elif typename in {"Worksheets", "Sheets"}:
            return xlWorksheetsSheetsWrapper(retval)
        elif typename == "Worksheet":
            return xlWorksheetWrapper(retval)
        elif typename == "Range":
            return xlRangeWrapper(retval)
        else:
            if hasattr(retval, "_prop_map_get_"):
                if retval._prop_map_get_.get("Count", False):
                    # here not yet wrapped Collection objects are identified
                    retval = UtilsOffice.create_msoCollectionWrapper(retval)
                elif retval._prop_map_get_.get("Application", False):
                    # here other objects are identified
                    retval = UtilsOffice.create_msoBaseWrapper(retval)
                #  redirect result wrapper finder to Excel result wrapper
                retval._wrap_retval = xlGenericWrapper._wrap_retval
            return retval

    # pass through for Microsoft Office object model set-attributes to wrapped COM object
    def __setattr__(self, attr, value) -> Any:

        # setattr for "pythonized" caller names - private wrapper attributes and
        # wrapped object methods for set
        # old assumption: wrapper class has no attributes to be set externally / properties
        # but re-activated
        attr_lower = attr.lower()
        attr_desnaked = attr_lower.replace("_", "")
        if attr.lower() in self.__class__._cls_attrmap:
            super().__setattr__(self.__class__._cls_attrmap[attr_lower], value)
        elif attr_desnaked in self.__class__._cls_attrmap:
            super().__setattr__(self.__class__._cls_attrmap[attr_desnaked], value)
        if attr == "_xlWrapped" or attr[0:1] == "_":
            super().__setattr__(attr, value)
        elif attr_lower in self.__class__._cls_attrmap_wrapped_put:
            setattr(self._xlWrapped, self.__class__._cls_attrmap_wrapped_put[attr_lower], value)
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_put:
            setattr(self._xlWrapped, self.__class__._cls_attrmap_wrapped_put[attr_desnaked], value)
        # setattr must consider own class but not call hasattr / __getattr__ (avoid recursion)
        elif attr in self.__class__._cls_attrmap_wrapped_put.values():  # equivalent to hasattr(self._xlWrapped, attr)
            # __setattr__ must not call hasattr / __getattr__ (avoid recursion)
            setattr(self._xlWrapped, attr, value)
        else:
            err_msg = f"'{self!r}' object has no attribute '{attr}'"
            raise AttributeError(err_msg)


# forward declarations for ruff
xlWorkbookWrapper = NewType("xlWorkbookWrapper", object)
xlRangeWrapper = NewType("xlRangeWrapper", object)


# xlAppWrapper to extend standard Excel Application COM object and encapsulate extended functions

class xlAppWrapper:
    """
    xlAppWrapper - wrapper class for Excel application object

    Singelton Excel app object is enforced via inner wrapper class.
    Note that "classic" singleton approach via __new__ does not work as
    __init__ must not have a parameter as required in this case.

    Note:
    Wrapper methods are named according to camel case naming convention like wrapped Excel object.
    However, via generic wrapper functionality the 'pythonic' access is ensured.
    """

    _innerWrapper: ClassVar[object] = None

    def __init__(self, xlWrapped=None):  # docsig: disable=SIG102

        if not xlAppWrapper._innerWrapper:
            xlAppWrapper._innerWrapper = xlAppWrapper.__xlAppWrapper(xlWrapped)
        else:
            if xlWrapped is not None:
                # here identity of application object could be checked
                # xlAppWrapper._innerWrapper._xlWrapped = xlWrapped
                pass

    def __getitem__(self, idx: Union[int, str]):
        return self._innerWrapper.__getitem__(idx)  # type: ignore

    def __getattr__(self, attr):
        return getattr(self._innerWrapper, attr) if attr != "_innerWrapper" else getattr(self, attr)

    def __setattr__(self, attr, value):
        setattr(self._innerWrapper, attr, value) if attr != "_innerWrapper" else super().__setattr__(attr, value)

    class __xlAppWrapper(xlGenericWrapper):
        """
        __xlAppWrapper - inner wrapper class for Excel application object
        """

        def __init__(self, xlWrapped=None):  # docsig: disable=SIG102

            b_started = False
            if xlWrapped is None:
                xlWrapped, b_started = UtilsOffice.assignCOMapplication(UtilsOffice.COMclass_Excel, True)
            super().__init__(xlWrapped)
            self._ExcelFlags = ExcelFlagsClass()
            self._started = b_started
            self._opened_workbooks_when_started = [wb.FullName for wb in self._xlWrapped.Workbooks]
            self._ExcelFlags.initialized = False
            atexit.register(self.CleanUpAndQuit)

        def __getitem__(self, idx: Union[int, str]):
            return xlWorkbookWrapper(self._xlWrapped.Workbooks(idx))

        def saveExcelFlags(self) -> None:
            """
            saveExcelFlags - wrapper object method calling function to save Excel flags
            """
            saveExcelFlags(self._xlWrapped, self._ExcelFlags)

        def setExcelFlags(self, call_level: str = "", force_save: bool = False, **kwargs) -> None:  # docsig: disable=SIG203
            """
            setExcelFlags - wrapper object method calling function to set Excel flags

            Args:
                call_level (str, optional): call level. save before setting if call_level is 'main'. Defaults to "".
                force_save (bool, optional): flag to enforce save before setting Excel flags. Defaults to False.
                **kwargs (_type_): Excel flags to be set as keyword parameters
            """
            if (not self._ExcelFlags.initialized) or (call_level.find("main") > 0) or force_save:
                self.saveExcelFlags()
            setExcelFlags(self._xlWrapped, **kwargs)

        def resetExcelFlags(self) -> None:
            """
            resetExcelFlags - wrapper object method calling function to reset Excel flags
            """
            resetExcelFlags(self._xlWrapped, self._ExcelFlags)

        def isWorkbookOpen(self, wbname: str) -> bool:
            """
            isWorkbookOpen - wrapper object method calling function to check if workbook is open

            Note:
            Following Excel object hierarchy method assignment this method would be
            assigned to Workbooks level but assignment to application level seems
            adequate.

            Args:
                wbname (str): workbook name

            Returns:
                bool: boolean value indicating if workbook with name 'wbname' is open
            """
            return isWorkbookOpen(self._xlWrapped, wbname)

        def isWorkbookOpenFullname(self, wbname: str) -> bool:
            """
            isWorkbookOpenFullname - wrapper object method calling function to check if workbook is open

            Note:
            Following Excel object hierarchy method assignment this method would be
            assigned to Workbooks level but assignment to application level seems
            adequate.

            Args:
                wbname (str): workbook name

            Returns:
                bool: boolean value indicating if workbook with fullname 'wbname' is open
            """
            return isWorkbookOpenFullname(self._xlWrapped, wbname)

        def openWorkbook(self, filename: str, *args: Any, **kwargs: Any) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
            """
            openWorkbook - wrapper object method calling function to open workbook

            Note:
            Following Excel object hierarchy method assignment this method would be
            assigned to Workbooks level but assignment to application level seems
            adequate.

            Args:
                filename (str): filename of workbook

            Returns:
                bool: boolean value indicating if 'filename' was opened as workbook
            """
            return openWorkbook(self._xlWrapped, filename, *args, **kwargs)

        def openText(self, filename: str, *args: Any, **kwargs: Any) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
            """
            openText - wrapper object method calling function to open workbook

            Note:
            Following Excel object hierarchy method assignment this method would be
            assigned to Workbooks level but assignment to application level seems
            adequate.

            Args:
                filename (str): filename of textfile

            Returns:
                bool: boolean value indicating if 'filename' was opened as textfile
            """
            return openText(self._xlWrapped, filename, *args, **kwargs)

        def close_workbooks_not_opened_when_started(self, *args: Any, **kwargs: Any) -> None:
            """
            close_workbooks_not_opened - close workbooks not open at startup
            """
            if not self._started:
                for wbname in [wb.FullName for wb in self._xlWrapped.Workbooks]:
                    if wbname not in self._opened_workbooks_when_started:
                        self._xlWrapped.Workbooks(os.path.basename(wbname)).Close(SaveChanges=False)

        def CleanUpAndQuit(self):
            """
            clean-up Excel application and quit if started by Python program
            """
            self.close_workbooks_not_opened_when_started()
            self.resetExcelFlags()
            if self._started:
                UtilsOffice.startedAppQuit(self._xlWrapped)


# xlAppWrapper methods as direct callables

def saveExcelFlags(xlapp: object, ExcelFlags: ExcelFlagsClass) -> None:
    """
    saveExcelFlags - save Excel flags a) Calculation b) EnableEvents and c) ScreenUpdating

    Args:
        xlapp (object): Excel application object
        ExcelFlags (ExcelFlagsClass): ExcelFlagClass object to save Excel flags
    """

    if not ExcelFlags.initialized:
        ExcelFlags.initialized = True
        if xlapp.Workbooks.Count > 0:
            ExcelFlags.Calculation = xlapp.Calculation
        ExcelFlags.EnableEvents = xlapp.EnableEvents
        ExcelFlags.ScreenUpdating = xlapp.ScreenUpdating

def save_excel_flags(xlapp: object, ExcelFlags: ExcelFlagsClass) -> None:
    """
    save_excel_flags - save Excel flags a) Calculation b) EnableEvents and c) ScreenUpdating

    Args:
        xlapp (object): Excel application object
        ExcelFlags (ExcelFlagsClass): ExcelFlagClass object to save Excel flags
    """
    saveExcelFlags(xlapp, ExcelFlags)


def setExcelFlags(xlapp: object, **kwargs) -> None:
    """
    setExcelFlags - set  Excel flags a) Calculation b) EnableEvents and c) ScreenUpdating

    Args:
        xlapp (object): Excel application object
        **kwargs (_type_): Excel flags to be set as keyword parameters
    """

    if xlapp.Workbooks.Count > 0:
        xlapp.Calculation = UtilsOffice.get_office_constant("xlCalculationManual")
    xlapp.EnableEvents = False
    xlapp.ScreenUpdating = False
    for key, value in kwargs.items():
        if key == "calculation":
            if xlapp.Workbooks.Count > 0:
                xlapp.Calculation = value
        if key == "enableevents":
            xlapp.EnableEvents = value
        if key == "screenupdating":
            xlapp.ScreenUpdating = value

def set_excel_flags(xlapp: object, **kwargs) -> None:
    """
    set_excel_flags - set  Excel flags a) Calculation b) EnableEvents and c) ScreenUpdating

    Args:
        xlapp (object): Excel application object
        **kwargs (_type_): Excel flags to be set as keyword parameters
    """
    setExcelFlags(xlapp, **kwargs)


def resetExcelFlags(xlapp: object, ExcelFlags: ExcelFlagsClass) -> None:
    """
    resetExcelFlags - reset/restore Excel flags from ExcelFlagClass object

    Args:
        xlapp (object): Excel application object
        ExcelFlags (ExcelFlagsClass): ExcelFlagClass object with saved Excel flags
    """

    if ExcelFlags.initialized:
        if xlapp.Workbooks.Count > 0:
            if ExcelFlags.Calculation != 0:
                xlapp.Calculation = ExcelFlags.Calculation
        xlapp.EnableEvents = ExcelFlags.EnableEvents
        xlapp.ScreenUpdating = ExcelFlags.ScreenUpdating

def reset_excel_flags(xlapp: object, ExcelFlags: ExcelFlagsClass) -> None:
    """
    reset_excel_flags - reset/restore Excel flags from ExcelFlagClass object

    Args:
        xlapp (object): Excel application object
        ExcelFlags (ExcelFlagsClass): ExcelFlagClass object with saved Excel flags
    """
    resetExcelFlags(xlapp, ExcelFlags)


def isWorkbookOpen(xlapp: object, wbname: str) -> bool:
    """
    isWorkbookOpen - check if workbook with name 'wbname' is open

    Args:
        xlapp (object): Excel application object
        wbname (str): name of workbook (name without path)

    Returns:
        bool: boolean value indicating if workbook with name 'wbname' is open
    """
    return wbname in [wb.Name for wb in xlapp.Workbooks]

def is_workbook_open(xlapp: object, wbname: str) -> bool:
    """
    is_workbook_open - check if workbook with name 'wbname' is open

    Args:
        xlapp (object): Excel application object
        wbname (str): name of workbook (name without path)

    Returns:
        bool: boolean value indicating if workbook with name 'wbname' is open
    """
    return wbname in [wb.Name for wb in xlapp.Workbooks]


# check if workbook is open in Excel - fullname including  path
def isWorkbookOpenFullname(xlapp: object, wbname: str) -> bool:
    """
    isWorkbookOpenFullname - check if workbook with name 'wbname' is open

    Args:
        xlapp (object): Excel application object
        wbname (str): fullname of workbook

    Returns:
        bool: boolean value indicating if workbook with fullname 'wbname' is open
    """
    return wbname in [wb.FullName for wb in xlapp.Workbooks]

def is_workbook_open_fullname(xlapp: object, wbname: str) -> bool:
    """
    is_workbook_open_fullname - check if workbook with name 'wbname' is open

    Args:
        xlapp (object): Excel application object
        wbname (str): fullname of workbook

    Returns:
        bool: boolean value indicating if workbook with fullname 'wbname' is open
    """
    return wbname in [wb.FullName for wb in xlapp.Workbooks]


def openWorkbook(xlapp: object, filename: str, minimizenew: bool = True, autoexec: bool = False, **kwargs) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
    """
    openWorkbook - open workbook

    Wrapper for opening workbook to allow processing of locked files.
    For parameters see Excel function signature.

    Args:
        xlapp (object): Excel application object
        filename (str): fullname of workbook
        minimizenew (bool, optional): flag if window for new file should be minimized. Defaults to True.
        autoexec (bool, optional): flag to surpress autoexec macro in opened workbook. Defaults to False.

    Returns:
        Union[xlWorkbookWrapper, None]: opened workbook or None
    """

    opened: bool = False

    if not os.path.isfile(filename):
    # if pathlib.Path(filename).is_file():
        err_msg = f"Trying to open '{filename}'. File does not exist."
        raise ErrorUtilsExcel(err_msg)
    basename: str = os.path.basename(filename)
    # basename = pathlib.Path(filename).name
    if isWorkbookOpenFullname(xlapp, filename):
        err_msg = f"Trying to open '{filename}'. File already open in Microsoft Excel."
        raise ErrorUtilsExcel(err_msg)
    if isWorkbookOpen(xlapp, basename):
        err_msg = f"Trying to open '{filename}'. Different file with same name '{basename}' already open in Microsoft Excel."
        raise ErrorUtilsExcel(err_msg)

    if xlapp.ActiveSheet is not None:
        activesheet: object = xlapp.ActiveSheet
        winstate: int = xlapp.ActiveWindow.WindowState
    else:
        activesheet = None
    excelflags = ExcelFlagsClass()
    saveExcelFlags(xlapp, excelflags)
    setExcelFlags(xlapp, EnableEvents=autoexec)

    # signature of Excel openWorkbook according to Excel object catalog
    paramsOpenWB = {
        "UpdateLinks": False, "ReadOnly": False, "Format": None, "Password": None, "WriteResPassword": None,
        "IgnoreReadOnlyRecommended": True, "Origin": None, "Delimiter": ";", "Editable": True, "Notify": True,
        "Converter": None, "AddToMru": False, "Local": False, "CorruptLoad": None
    }
    Utils.copydictfields(kwargs, paramsOpenWB)

    try:
        opened = True
        # does not work -> https://stackoverflow.com/questions/19450837/how-to-open-a-password-protected-excel-file-using-python
        # xlapp.Workbooks.Open(Filename=filename, UpdateLinks=UpdateLinks, Password=Password, AddToMru=False, **kwargs)  # try with provided password first
        xlapp.Workbooks.Open(filename, *[value for key, value in paramsOpenWB.items()])  # try with provided password first
    except Exception as ErrorOpenWorkbook:
        opened = False
        while True:
            flag = "OK"
            opened = True
            try:
                # xlapp.Workbooks.Open(filename, UpdateLinks, AddToMru=False, **kwargs)  # enforce password entry
                paramsOpenWB["Password"] = None
                xlapp.Workbooks.Open(filename, *[value for key, value in paramsOpenWB.items()])  # enforce password entry
            except Exception as ErrorOpenWorkbook:
                opened = False
                flag = simplegui.popup_ok_cancel(
                    f"Trying to open '{filename}'. Password is wrong. Try again or cancel?",
                    title="open Workbook"
                )
            finally:
                if opened or (flag == "Cancel"):
                    break  # noqa: B012

    if opened:
        retval = xlapp.ActiveWorkbook  # note: ActiveWorkbook should be automatically wrapped as xlWorkbookWrapper
        if minimizenew:
            minimizeWindows(xlapp.ActiveWorkbook)
        if activesheet is not None:
            activesheet.Activate()
            xlapp.ActiveWindow.WindowState = winstate
    else:
        retval = None
    resetExcelFlags(xlapp, excelflags)

    return retval

def open_workbook(xlapp: object, filename: str, minimizenew: bool = True, autoexec: bool = False, **kwargs) -> Union[xlWorkbookWrapper, None]:
    """
    open_workbook - open workbook

    Wrapper for opening workbook to allow processing of locked files.
    For parameters see Excel function signature.

     Args:
         xlapp (object): Excel application object
         filename (str): fullname of workbook
         minimizenew (bool, optional): flag if window for new file should be minimized. Defaults to True.
         autoexec (bool, optional): flag to surpress autoexec macro in opened workbook. Defaults to False.

     Returns:
         Union[xlWorkbookWrapper, None]: opened workbook or None
     """
    return openWorkbook(xlapp, filename, minimizenew, autoexec, **kwargs)


def openText(xlapp: object, filename: str, minimizenew: bool = True, **kwargs) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
    """
    openText - open text file

    Re-implemented OpenText method to overcome automatic recalculation bug in Excel.
    For parameters see Excel function signature.

    Note: openText initiates recalculation of open workbooks if not surpressed
    via appropriate setting of EnableCalculationFlag. The function disables
    the recalculating and resets the flag after loading the text file.

    Args:
        xlapp (object): Excel application object
        filename (str): filename of textfile
        minimizenew (bool, optional): flag to minimize new file. Defaults to True.

    Returns:
        bool: boolean value indicating if 'filename' was opened as textfile
    """

    opened = False

    if not os.path.isfile(filename):
    # if not pathlib.Path(filename).is_file:
        err_msg = f"File '{filename}' does not exist."
        raise ErrorUtilsExcel(err_msg)
    basename: str = os.path.basename(filename)
    # basename: str = pathlib.Path(filename).name
    if isWorkbookOpenFullname(xlapp, filename):
        err_msg = f"File '{filename}' is already open in Microsoft Excel."
        raise ErrorUtilsExcel(err_msg)
    if isWorkbookOpen(xlapp, basename):
        err_msg = f"Different file with same name '{basename}' is already open in Microsoft Excel."
        raise ErrorUtilsExcel(err_msg)

    if xlapp.ActiveSheet is not None:
        activesheet: object = xlapp.ActiveSheet
        winstate: int = xlapp.ActiveWindow.WindowState
    else:
        activesheet = None
    excelflags = ExcelFlagsClass()
    saveExcelFlags(xlapp, excelflags)
    setExcelFlags(xlapp)

    # set calculation flag off
    dictwb: dict = {}
    for i in range(1, xlapp.Workbooks.Count + 1):
        dictws: dict = {}
        for j in range(1, xlapp.Workbooks(i).Worksheets.Count + 1):
            dictws[xlapp.Workbooks(i).Worksheets(j).Name] = xlapp.Workbooks(i).Worksheets(j).EnableCalculation
            xlapp.Workbooks(i).Worksheets(j).EnableCalculation = False
        dictwb[xlapp.Workbooks(i).Name] = dictws

    # signature of Excel openText according to Excel object catalog
    paramsOpenText = {
        "Origin": None, "StartRow": 1, "DataType": None,
        "TextQualifier": UtilsOffice.get_office_constant("xlTextQualifierDoubleQuote"), "ConsecutiveDelimiter": False,
        "Tab": False, "Semicolon": False, "Comma": False, "Space": False, "Other": False, "OtherChar": None,
        "FieldInfo": None, "TextVisualLayout": None, "DecimalSeparator": ".", "ThousandsSeparator": ",",
        "TrailingMinusNumbers": False, "Local": False
    }
    Utils.copydictfields(kwargs, paramsOpenText)

    opened = True
    # xlapp.Workbooks.OpenText(Filename=Filename, *args, **kwargs)
    xlapp.Workbooks.OpenText(filename, *[value for key, value in paramsOpenText.items()])

    # reset calculation flag
    for wb, dictws in dictwb.items():
        for ws, calcflg in dictws.items():
            xlapp.Workbooks(wb).Worksheets(ws).EnableCalculation = calcflg

    if opened:
        retval = xlapp.ActiveWorkbook  # note: ActiveWorkbook should be automatically wrapped as xlWorkbookWrapper
        if minimizenew:
            minimizeWindows(xlapp.ActiveWorkbook)
        if activesheet is not None:
            activesheet.Activate()
            xlapp.ActiveWindow.WindowState = winstate
    else:
        retval = None
    resetExcelFlags(xlapp, excelflags)

    return retval

def open_text(xlapp: object, filename: str, minimizenew: bool = True, **kwargs) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
    """
    open_text - open text file

    Re-implemented OpenText method to overcome automatic recalculation bug in Excel.
    For parameters see Excel function signature.

    Note: openText initiates recalculation of open workbooks if not surpressed
    via appropriate setting of EnableCalculationFlag. The function disables
    the recalculating and resets the flag after loading the text file.

    Args:
        xlapp (object): Excel application object
        filename (str): filename of textfile
        minimizenew (bool, optional): flag to minimize new file. Defaults to True.

    Returns:
        bool: boolean value indicating if 'filename' was opened as textfile
    """
    return openText(xlapp, filename, minimizenew, **kwargs)


# xlWorkbooksWrapper to extend standard Excel Workbook COM object and encapsulate extended functions

class xlWorkbooksWrapper(xlGenericWrapper):
    """
    xlWorkbooksWrapper - wrapper for collection objects with pass through for direct calls to wrapped object

    Workbooks wrapper is same generic office wrapper but needed to ensure return type Workbook.

    Note:
    Wrapper methods are named according to camel case naming convention like wrapped Excel object.
    However, via generic wrapper functionality the 'pythonic' access is ensured.
    """

    def __call__(self, idx: Union[int, str]):

        if isinstance(idx, (int, str)):
            return xlWorkbookWrapper(self._xlWrapped.Item(idx))
        else:
            return self._xlWrapped

    def __getitem__(self, idx: Union[int, str]):

        if isinstance(idx, (int, str)):
            return xlWorkbookWrapper(self._xlWrapped.Item(idx))
        else:
            return self._xlWrapped

    # object-specific iterator class to support list-comprehension even for different index bases in Python and COM
    class _xlWorkbooksIterator(UtilsOffice.GenericIterator):

        def __next__(self):
            self.index += 1
            if self.index <= self.data.Count:
                return xlWorkbookWrapper(self.data(self.index))
            else:
                raise StopIteration

    def __iter__(self):
        return self._xlWorkbooksIterator(self._xlWrapped)

    def isWorkbookOpen(self, wbname: str) -> bool:
        """
        isWorkbookOpen - wrapper object method calling function to check if workbook is open

        Args:
            wbname (str): workbook name

        Returns:
            bool: boolean value indicating if workbook with name 'wbname' is open
        """
        return xlAppWrapper(self._xlWrapped.Parent).isWorkbookOpen(wbname)

    def isWorkbookOpenFullname(self, wbname: str) -> bool:
        """
        isWorkbookOpenFullname - wrapper object method calling function to check if workbook is open

        Args:
            wbname (str): workbook name

        Returns:
            bool: boolean value indicating if workbook with fullname 'wbname' is open
        """
        return xlAppWrapper(self._xlWrapped.Parent).isWorkbookOpenFullname(wbname)

    def openWorkbook(self, filename: str, *args: Any, **kwargs: Any) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
        """
        openWorkbook - wrapper object method calling function to open workbook

        Args:
            filename (str): filename of workbook

        Returns:
            bool: boolean value indicating if 'filename' was opened as workbook
        """
        return xlAppWrapper(self._xlWrapped.Parent).openWorkbook(filename, *args, **kwargs)

    def openText(self, filename: str, *args: Any, **kwargs: Any) -> Union[xlWorkbookWrapper, None]:  # docsig: disable=SIG203
        """
        openText - wrapper object method calling function to open workbook

        Args:
            filename (str): filename of textfile

        Returns:
            bool: boolean value indicating if 'filename' was opened as textfile
        """
        return xlAppWrapper(self._xlWrapped.Parent).openText(filename, *args, **kwargs)


# xlWorkbookWrapper to extend standard Excel Workbook COM object and encapsulate extended functions

class xlWorkbookWrapper(xlGenericWrapper):  # type: ignore[no-redef]
    """
    xlWorkbookWrapper - wrapper class for Excel workbook object

    Note:
    Wrapper methods are named according to camel case naming convention like wrapped Excel object.
    However, via generic wrapper functionality the 'pythonic' access is ensured.
    """

    def __eq__(self, other: object) -> bool:
        """
        __eq__ - check if two workbook objects are the same via Excel object attributes
        """
        if not isinstance(other, xlWorkbookWrapper):  # type: ignore
            return NotImplemented
        return self._xlwrapped.FullName == other._xlwrapped.FullName

    def __hash__(self):
        return hash(self._xlwrapped.FullName)

    def __getitem__(self, idx: Union[int, str]):
        return xlWorksheetWrapper(self._xlWrapped.Worksheets(idx))

    def sheet_exists(self, wsname: str) -> bool:
        """
        sheet_exists - wrapper object method calling function to check if worksheet 'wsname' exists

        Args:
            wsname (str): name of worksheet

        Returns:
            bool: flag if worksheet 'wsname' exist in workbook
        """
        return sheet_exists(self._xlWrapped, wsname)

    def containsWorksheet(self, wsname: str) -> bool:
        """
        containsWorksheet - wrapper object method calling function to check if worksheet 'wsname' exists

        Args:
            wsname (str): name of worksheet

        Returns:
            bool: flag if worksheet 'wsname' exist in workbook
        """
        return self.sheet_exists(wsname)

    def deleteWorksheet(self, wsname: str) -> None:
        """
        deleteWorksheet - wrapper object method calling function to delete worksheet 'wsname' from workbook

        Args:
            wsname (str): name of worksheet
        """
        deleteWorksheet(self._xlWrapped, wsname)

    def minimizeWindows(self) -> None:
        """
        minimizeWindows - wrapper object method calling function to minimize windows of workbook
        """
        minimizeWindows(self._xlWrapped)


# xlWorkbookWrapper methods as direct callables

def sheet_exists(workbook: object, wsname: str) -> bool:
    """
    sheet_exists - check if worksheet 'wsname' exists in workbook

    Args:
        workbook (object): workbook object
        wsname (str): name of worksheet

    Returns:
        bool: flag if worksheet 'wsname' exist in workbook
    """
    return wsname in [ws.Name for ws in workbook.Worksheets]

def containsWorksheet(workbook: object, wsname: str) -> bool:
    """
    containsWorksheet - check if worksheet 'wsname' exists in workbook

    Args:
        workbook (object): workbook object
        wsname (str): name of worksheet

    Returns:
        bool: flag if worksheet 'wsname' exist in workbook
    """
    return sheet_exists(workbook, wsname)

def contains_worksheet(workbook: object, wsname: str) -> bool:
    """
    contains_Worksheet - check if worksheet 'wsname' exists in workbook

    Args:
        workbook (object): workbook object
        wsname (str): name of worksheet

    Returns:
        bool: flag if worksheet 'wsname' exist in workbook
    """
    return sheet_exists(workbook, wsname)


def deleteWorksheet(workbook: object, wsname: str) -> None:
    """
    deleteWorksheet - delete worksheet 'wsname' from workbook

    Args:
        workbook (object): workbook object
        wsname (str): name of worksheet
    """

    if containsWorksheet(workbook, wsname):
        savedDisplayAlerts: bool = workbook.Parent.DisplayAlerts
        workbook.Parent.DisplayAlerts = False
        workbook.Worksheets(wsname).Delete()
        workbook.Parent.DisplayAlerts = savedDisplayAlerts

def delete_worksheet(workbook: object, wsname: str) -> None:
    """
    delete_worksheet - delete worksheet 'wsname' from workbook

    Args:
        workbook (object): workbook object
        wsname (str): name of worksheet
    """
    deleteWorksheet(workbook, wsname)


def minimizeWindows(workbook: object) -> None:
    """
    minimizeWindows - minimize windows of workbook

    Args:
        workbook (object): workbook object
    """

    for window in [1, workbook.Windows.Count]:
        workbook.Windows(window).WindowState = UtilsOffice.get_office_constant("xlMinimized")

def minimize_windows(workbook: object) -> None:
    """
    minimize_windows - minimize windows of workbook

    Args:
        workbook (object): workbook object
    """
    minimizeWindows(workbook)


# xlWorksheetsWrapper to extend standard Excel Worksheet COM object and encapsulate extended functions

class xlWorksheetsSheetsWrapper(xlGenericWrapper):
    """
    xlWorksheetsSheetsWrapper - wrapper for collection objects with pass through for direct calls to wrapped object

    Worksheets wrapper is same like the generic office wrapper but needed to ensure return type Worksheet.
    """

    def __call__(self, idx: Union[int, str]):

        if isinstance(idx, (int, str)):
            typename = type(self._xlWrapped.Item(idx)).__name__.replace("_", "")
            if typename == "Worksheet":
                return xlWorksheetWrapper(self._xlWrapped.Item(idx))
            elif typename in {"Chart", "DialogSheet"}:
                return UtilsOffice.msoBaseWrapper(self._xlWrapped.Item(idx))
        else:
            return self._xlWrapped

    def __getitem__(self, idx: Union[int, str]):

        if isinstance(idx, (int, str)):
            typename = type(self._xlWrapped.Item(idx)).__name__.replace("_", "")
            if typename == "Worksheet":
                return xlWorksheetWrapper(self._xlWrapped.Item(idx))
            elif typename in {"Chart", "DialogSheet"}:
                return UtilsOffice.msoBaseWrapper(self._xlWrapped.Item(idx))
            else:
                return self._xlWrapped

    # object-specific iterator class to support list-comprehension even for different index bases in Python and COM
    class _xlWorksheetsSheetsIterator(UtilsOffice.GenericIterator):

        def __next__(self):
            self.index += 1
            if self.index <= self.data.Count:
                retval = self.data(self.index)
                typename = type(retval).__name__.replace("_", "")
                if typename == "Worksheet":
                    return xlWorksheetWrapper(retval)
                elif typename in {"Chart", "DialogSheet"}:
                    return UtilsOffice.msoBaseWrapper(retval)
            else:
                raise StopIteration

    def __iter__(self):
        return self._xlWorksheetsSheetsIterator(self._xlWrapped)


# xlWorksheetWrapper to extend standard Excel Worksheet COM object and encapsulate extended functions

class xlWorksheetWrapper(xlGenericWrapper):  # noqa: PLW1641
    """
    xlWorksheetWrapper - wrapper class for Excel worksheet object

    Note:
    Wrapper methods are named according to camel case naming convention like wrapped Excel object.
    However, via generic wrapper functionality the 'pythonic' access is ensured.
    """

    def __eq__(self, other: object) -> bool:
        """
        __eq__ - check if two worksheet objects are the same via Excel object attributes
        """
        if not isinstance(other, xlWorksheetWrapper):
            return NotImplemented
        return \
            (self._xlwrapped.Parent.FullName == other._xlwrapped.Parent.FullName) and \
            (self._xlwrapped.Name == other._xlwrapped.Name)

    def __getitem__(self, *args):
        """
        __getitem__ - read cell directly via address in Excel A1 or R1C1 notation
        """
        if len(args) == 0:
            return self
        else:
            if isinstance(*args, tuple):
                return self.Range(args[0][0], args[0][1])
            else:
                return self.Range(*args)

    def __setitem__(self, *args):
        """
        __setitem__ - write cell directly via address in Excel A1 or R1C1 notation
        """
        if len(args) == 2:  # noqa: PLR2004
            if isinstance(args[0], str):
                tmp_range = self._xlWrapped.Range(args[0])
                tmp_range.Value = args[1]
            elif isinstance(args[0], tuple):
                tmp_range = self._xlWrapped.Cells(args[0][0], args[0][1])
                tmp_range.Value = args[1]
        else:
            raise RuntimeError

    def Range(self, *range: list[Union[str, tuple[int, int], tuple[tuple[int, int], tuple[int, int]]]]) -> xlRangeWrapper:  # docsig: disable=SIG203
        """
        Range - create Range object for provided identifier

        Args:
            *range (list[Union[str, tuple[int, int], tuple[tuple[int, int], tuple[int, int]]]]): range identifier

        Returns:
            xlRangeWrapper: wrapped range object
        """

        if isinstance(range[0], str) and len(range) == 1:
            return xlRangeWrapper(self._xlWrapped.Range(range[0]))
        elif isinstance(range[0], int):
            if len(range) == 2:  # noqa: PLR2004
                if isinstance(range[1], int):
                    return xlRangeWrapper(self._xlWrapped.Cells(range[0], range[1]))
        elif isinstance(range[0], tuple):
            if len(range) == 2:  # noqa: PLR2004
                return xlRangeWrapper(
                    self._xlWrapped.Range(
                        self._xlWrapped.Cells(range[0][0], range[0][1]),
                        self._xlWrapped.Cells(range[1][0], range[1][1])
                    )
                )
        raise RuntimeError

    # necessary as Columns is Excel internally a range object -> enforce resolution in Excel
    def Columns(self, column: int) -> xlRangeWrapper:
        """
        Columns - create range object for worksheet column

        Args:
            column (int): column number

        Returns:
            xlRangeWrapper: wrapped range object
        """
        return xlRangeWrapper(self._xlWrapped.Columns(column))

    # necessary as Rows is Excel internally a range object -> enforce resolution in Excel
    def Rows(self, row: int) -> xlRangeWrapper:
        """
        Rows - create range object for worksheet row

        Args:
            row (int): row number

        Returns:
            xlRangeWrapper: wrapped range object
        """
        return xlRangeWrapper(self._xlWrapped.Rows(row))

    def lastfilledRow(self, check_cols: Optional[list[int]] = None, rowstart: int = 1) -> int:
        """
        lastfilledRow - wrapper object method calling function to determine last filled row

        Args:
            check_cols (Optional[list[int]], optional): cols to be checked. Defaults to None.
            rowstart (int, optional): optional start row. Defaults to 1.

        Returns:
            int: last filled row with 'check_cols' not empty
        """
        if check_cols is None:
            check_cols = [1]
        return lastfilledRow(self._xlWrapped, check_cols, rowstart)

    def exportWorksheet(self, filename: str = "", format: str = ".csv") -> None:
        """
        exportWorksheet - wrapper object method calling function to export worksheet

        Args:
            filename (str, optional): file name for worksheet export. Defaults to "".
            format (str, optional): file format for worksheet export. Defaults to ".csv".
        """
        exportWorksheet(self._xlWrapped, filename)


# xlWorksheetWrapper methods as direct callables

def lastfilledRow(ws: object, check_cols: Optional[list[int]] = None, rowstart: int = 1) -> int:
    """
    lastfilledRow - determine last filled row of worksheet

    Args:
        ws (object): worksheet object
        check_cols (List[int] or None], optional): cols to be checked. Defaults to None.
        rowstart (int, optional): optional start row. Defaults to 1.

    Returns:
        int: last filled row with 'check_cols' not empty
    """

    if check_cols is None:
        check_cols = [1]

    # note: order of parameters important, no named parameters calling Excel !
    # checkrow = ws.Cells.Find(What="*", After=ws.Cells(1, 1), LookIn=xlValues, LookAt=xlPart, SearchOrder=xlByRows, SearchDirection=xlPrevious).Row
    checkrow: int = ws.Cells.Find("*", ws.Cells(1, 1),
        UtilsOffice.get_office_constant("xlValues"), UtilsOffice.get_office_constant("xlPart"),
        UtilsOffice.get_office_constant("xlByRows"), UtilsOffice.get_office_constant("xlPrevious")
    ).Row

    # check defined columns -> mitigation for malfunction
    check = True
    for col in check_cols:
        check = check and (ws.Cells(checkrow + 1, col).Value is None)
    if not check:
        err_msg = "Error determining last filled row."
        raise ErrorUtilsExcel(err_msg)

    return checkrow

def last_filled_row(ws: object, check_cols: Optional[list[int]] = None, rowstart: int = 1) -> int:
    """
    last_filled_row - determine last filled row of worksheet

    Args:
        ws (object): worksheet object
        check_cols (List[int] or None], optional): cols to be checked. Defaults to None.
        rowstart (int, optional): optional start row. Defaults to 1.

    Returns:
        int: last filled row with 'check_cols' not empty
    """
    return lastfilledRow(ws, check_cols, rowstart)


def exportWorksheet(ws: object, filename: str = "", format: str = ".csv") -> None:
    """
    exportWorksheet - export single worksheet as file

    Args:
        ws (object): worksheet object for export
        filename (str, optional): file name for worksheet export. Defaults to "".
        format (str, optional): file format for worksheet export. Defaults to ".csv".
    """

    checkxlFileFormat(format)
    parampath = os.path.dirname(filename)
    # parampath = pathlib.path(filename).parent
    if parampath == "":
        parampath = tempfile.gettempdir()
    parambase, paramext = os.path.splitext(os.path.basename(filename))
    # parambase = pathlib.Path(filename).stem
    if parambase == "":
        parambase, paramext = os.path.splitext(ws.Parent.Name)
        parambase = parambase + "_" + ws.Name
        # parambase = pathlib.Path(ws.Parent.Name).stem
    exportfile = os.path.join(parampath, parambase + format)
    # exportfile = parampath.joinpath(parambase + format)
    if os.path.isfile(exportfile):
    # if pathlib.Path(exportfile).is_file():
        err_msg = f"File '{exportfile}' does already exist."
        raise ErrorUtilsExcel(err_msg)

    try:
        ws.Copy(Before=None, After=None)  # defining Before and After explicitly necessary to achieve org. Excel behaviour
        xlapp = ws.Parent.Parent
        xlapp.ActiveWorkbook.SaveAs(exportfile, extcodes[format])
        xlapp.ActiveWorkbook.Close()
    except pythoncom.com_error:
        err_msg = f"Failed to save sheet '{ws.Name}' as '{exportfile}'."
        raise ErrorUtilsExcel(err_msg)  # noqa: B904

def export_worksheet(ws: object, filename: str = "", format: str = ".csv") -> None:
    """
    export_worksheet - export single worksheet as file

    Args:
        ws (object): worksheet object for export
        filename (str, optional): file name for worksheet export. Defaults to "".
        format (str, optional): file format for worksheet export. Defaults to ".csv".
    """
    exportWorksheet(ws, filename, format)


# xlRangeWrapper to extend standard Excel Range COM object and encapsulate extended functions

class xlRangeWrapper(xlGenericWrapper):  # type: ignore[no-redef]
    """
    xlRangeWrapper - wrapper class for Excel range object

    Note:
    Wrapper methods are named according to camel case naming convention like wrapped Excel object.
    However, via generic wrapper functionality the 'pythonic' access is ensured.
    """

    def __eq__(self, other: object) -> bool:
        """
        __eq__ - check if two range objects are the same via Excel object attributes
        """
        if not isinstance(other, xlRangeWrapper):  # type: ignore[misc]
            return NotImplemented
        return \
            (self._xlwrapped.Parent.Parent.FullName == other._xlwrapped.Parent.Parent.FullName) and \
            (self._xlwrapped.Parent.FullName == other._xlwrapped.Parent.Name) and \
            ((self._xlwrapped.Name == other._xlwrapped.Name) or (self._xlwrapped.Address == other._xlwrapped.Address))

    def __hash__(self):
        return hash(self._xlwrapped.Parent.Parent.FullName + self._xlwrapped.Parent.FullName + self._xlwrapped.Name + self._xlwrapped.Address)

    @multimethod.multimethod
    def __call__(self):
        return self

    @multimethod.multimethod   # type: ignore
    def __call__(self, range: str):  # noqa: F811
        return xlRangeWrapper(self._xlWrapped.Range(range))

    @multimethod.multimethod  # type: ignore
    def __call__(self, cell: tuple[int, int]):  # noqa: F811
        return xlRangeWrapper(self._xlWrapped.Cells(cell[0], cell[1]))

    @multimethod.multimethod  # type: ignore
    def __call__(self, row: int, col: int):  # noqa: F811
        return xlRangeWrapper(self._xlWrapped.Cells(row, col))

    @multimethod.multimethod  # type: ignore
    def __call__(self, corner1: tuple[int, int], corner2: tuple[int, int]):  # noqa: F811
        return xlRangeWrapper(
            self._xlWrapped.Range(
                self._xlWrapped.Cells(corner1[0], corner1[1]),
                self._xlWrapped.Cells(corner2[0], corner2[1])
            )
        )

    def __getitem__(self, *args):

        if len(args) == 0:
            return self
        elif isinstance(args[0], str) and len(range) == 1:
            return xlRangeWrapper(self._xlWrapped.Range(args[0]))
        elif isinstance(args[0], tuple):
            if len(args) == 1 and len(args[0]) == 2:  # noqa: PLR2004
                if isinstance(args[0][0], int) and isinstance(args[0][1], int):
                    return xlRangeWrapper(self._xlWrapped.Cells(args[0][0], args[0][1]))
        raise RuntimeError

    def Range(self, *range: list[Union[str, tuple[int, int], tuple[tuple[int, int], tuple[int, int]]]]) -> xlRangeWrapper:  # docsig disable=SIG203
        """
        Range - create Range object for provided identifier

        Args:
            *range (list[Union[str, tuple[int, int], tuple[tuple[int, int], tuple[int, int]]]]): range identifier

        Returns:
            xlRangeWrapper: wrapped range object
        """

        if isinstance(range[0], str) and len(range) == 1:
            return xlRangeWrapper(self._xlWrapped.Range(range[0]))
        elif isinstance(range[0], int):
            if len(range) == 2:  # noqa: PLR2004
                if isinstance(range[1], int):
                    return xlRangeWrapper(self._xlWrapped.Cells(range[0], range[1]))
        elif isinstance(range[0], tuple):
            if len(range) == 2:  # noqa: PLR2004
                return xlRangeWrapper(
                    self._xlWrapped.Range(
                        self._xlWrapped.Cells(range[0][0], range[0][1]),
                        self._xlWrapped.Cells(range[1][0], range[1][1])
                    )
                )
        raise RuntimeError

    def Sort(self, *args: Any, **kwargs: Any):
        """
        sort - wrapper object method calling sort method for range
        """
        sortRange(self._xlWrapped, *args, **kwargs)

    def setDate(self, datevalue: datetime.datetime) -> None:
        """
        setDate - wrapper object method calling function to set date overcoming timezone==None issue
        """
        setDate(self._xlWrapped, datevalue)

    def Dims(self) -> tuple:
        """
        Dims - return dimensions of range object
        """
        return self._xlWrapped.Rows.Count, self._xlWrapped.Columns.Count

    def Range2PandasDF(self) -> pandas.DataFrame:
        """
        Range2PandasDF - copy range to dataframe (without header)

        Returns:
            pandas.DataFrame: data frame with range values
        """
        return pandas.DataFrame(list(self._xlWrapped.Value))

    def Range2PandasDFheader(self) -> pandas.DataFrame:
        """
        Range2PandasDFheader - copy range to dataframe using first row as header line

        Returns:
            pandas.DataFrame: data frame with range values
        """
        return pandas.DataFrame(self._xlWrapped.Value.Value[1:], columns=self._xlWrapped.Value.Rows(1).Value[0])

    def Values2Range(self, values, autoadjust: bool = False, header: bool = True) -> None:
        """
        values2range - wrapper object method calling function to copy values to Excel range object

        Args:
            values (_type_): source values to be copied to target range
            autoadjust (bool, optional): Flag to control adjustment of target range according to source data. Defaults to False.
            header (bool, optional): Flag if header form source dataframe should be transferred. Defaults to True.
        """
        values2range(self._xlWrapped, values, autoadjust, header)


# xlRangeWrapper methods as direct callables

def sortRange(xlRange: object, *args: Any, **kwargs: Any) -> None:  # docsig: disable=SIG302
    """
    sortRange - sort for Excel range object (for parameters see Excel function signature)

    Args:
        xlRange (object): Excel range object
        *args (Any): n. a.
        **kwargs (Any): parameters according to sort parameters in Excel object catalog
    """

    # signature of Excel sort according to Excel object catalog
    paramsSortRange = {
        "Key1": None, "Order1": UtilsOffice.get_office_constant("xlAscending"),
        "Key2": None, "Type": None, "Order2": UtilsOffice.get_office_constant("xlAscending"),
        "Key3": None, "Order3": UtilsOffice.get_office_constant("xlAscending"),
        "Header": UtilsOffice.get_office_constant("xlNo"), "OrderCustom": None, "MatchCase": True,
        "Orientation": UtilsOffice.get_office_constant("xlSortRows"),
        "SortMethod": UtilsOffice.get_office_constant("xlPinYin"),
        "DataOption1": UtilsOffice.get_office_constant("xlSortNormal"),
        "DataOption2": UtilsOffice.get_office_constant("xlSortNormal"),
        "DataOption3": UtilsOffice.get_office_constant("xlSortNormal")
    }
    Utils.copydictfields(kwargs, paramsSortRange)

    xlRange.Sort(*[value for key, value in paramsSortRange.items()])

def sort_range(xlRange: object, *args: Any, **kwargs: Any) -> None:  # docsig: disable=SIG302
    """
    sort_range - sort for Excel range object (for parameters see Excel function signature)

    Args:
        xlRange (object): Excel range object
        *args (Any): n. a.
        **kwargs (Any): parameters according to sort parameters in Excel object catalog
    """
    sortRange(xlRange, args, **kwargs)


def setDate(xlcell: object, datevalue: datetime.datetime) -> None:
    """
    setDate - set date from datetime object overcoming timezone==None issue

    Problem arises if a datetime object without timezone info (i. e. from
    strptime) is assigned to an Excel cell. The result of strptime timezone
    might be None but when assigned to Excel, timezone information is added
    automatically and may lead to a wrong date in Excel depending on timezone
    difference to UTC.

    Args:
        xlcell (object): target Excel cell
        datevalue (datetime.datetime): datetime value without timezone information
    """

    if xlcell.Rows.Count != 1 or xlcell.Columns.Count != 1:
        err_msg = "Error calling Excel date setter - only single cells allowed."
        raise ErrorUtilsExcel(err_msg)

    if isinstance(datevalue, datetime.date):
        datevalue = datetime.datetime.combine(datevalue, datetime.time.min)

    datevalue = datevalue.replace(microsecond=0) + datetime.timedelta(seconds=round(datevalue.microsecond / 10 ** 6))
    xlcell.Value = datevalue
    if xlcell.Value != datevalue:
        datevalue_tz_replace = datevalue.replace(tzinfo=xlcell.Value.tzinfo)
        xlcell.Value = datevalue_tz_replace
    if xlcell.Value.replace(tzinfo=None) != datevalue:
        err_msg = "Error setting date in Excel."
        raise ErrorUtilsExcel(err_msg)

def set_date(xlcell: object, datevalue: datetime.datetime) -> None:
    """
    set_date - set date from datetime object overcoming timezone==None issue

    Args:
        xlcell (object): target Excel cell
        datevalue (datetime.datetime): datetime value without timezone information
    """
    setDate(xlcell, datevalue)


def values2range(range: object, values, autoadjust: bool = False, header: bool = True) -> None:
    """
    values2range - copy values to Excel range object

    Args:
        range (object): target range
        values (_type_): source values to be copied to target range
        autoadjust (bool, optional): flag to control adjustment of target range according to source data. Defaults to False.
        header (bool, optional): flag if header form source dataframe should be transferred. Defaults to True.
    """

    def is_iter(value) -> bool:
        return hasattr(value, '__iter__') and not isinstance(value, str)

    # prepare dataframe
    if isinstance(values, pandas.core.frame.DataFrame):
        if header:
            colsheader = tuple(col for col in values.columns)
        values = tuple(map(tuple, values.values))
        if header:
            values = (colsheader,) + values  # noqa: RUF005

    rows: int = range.Rows.Count
    cols: int = range.Columns.Count

    # reshape values
    if not is_iter(values):
        values = (values,)
    elif any(is_iter(v) for v in values):
        values = tuple(v if is_iter(v) else (v,) for v in values)
    else:
        values = (values,)
    # auto-adjust target range
    rows_values = len(values)
    cols_values = max([len(v) if is_iter(v) else 1 for v in values])  # noqa: C419
    if not autoadjust:
        if rows_values > rows or cols_values > cols:
            err_msg = "Dimensions of values passed exceed dimensions of target range object."
            raise ErrorUtilsExcel(err_msg)
    else:
        rows = rows_values
        cols = cols_values
        range.Cells(1, 1).Resize(rows, cols)
    array = numpy.full(shape=(rows, cols), fill_value=None)
    # fill empty values - as id-check for numpy.nan does not work reliably, do double-check
    for i, value in enumerate(values):
        if isinstance(value, (list, tuple)):
            for j, v in enumerate(value):
                array[i, j] = v if (id(v) != id(numpy.nan)) and (str(v).lower() != "nan") else None
        else:
            array[0, i] = value if (id(value) != id(numpy.nan)) and (str(v).lower() != "nan") else None
    values = tuple(map(tuple, array))

    range.Value2 = values



# other stuff


# code for supporting filetype check
extcodes = {
    '.xla': 18, '.csv': 6, '.txt': -4158, '.dif': 9, '.xlsb': 50, '.htm': 44, '.html': 44, '.ods': 60,
    '.xlam': 55, '.xltx': 54, '.xltm': 53, '.xlsx': 51, '.xlsm': 52, '.xlt': 17, '.xls': -4143, '.xml': 46,
}

def checkxlFileFormat(filename: str, raiseerror: bool = True) -> bool:
    """
    checkxlFileFormat - check file format support by Excel based on extension.

    Args:
        filename (str): File name
        raiseerror (bool, optional): control raising error if file format not valid. Defaults to True.

    Returns:
        bool: file format of 'filename' is supported by Excel
    """

    parambase, paramext = os.path.splitext(os.path.basename(filename))
    # parambase = pathlib.Path(filename).stem
    # paramext = pathlib.Path(filename).suffix
    if (paramext is not None) and (paramext != ""):
        if paramext not in extcodes:
            if raiseerror:
                err_msg = f"Filetype '{paramext}' is not a supported format for Microsoft Excel."
                raise ErrorUtilsExcel(err_msg)
            else:
                return False
        else:
            return True

    return False

def check_xl_fileformat(filename: str, raiseerror: bool = True) -> bool:
    """
    check_xl_fileformat - check file format support by Excel based on extension.

    Args:
        filename (str): File name
        raiseerror (bool, optional): control raising error if file format not valid. Defaults to True.

    Returns:
        bool: file format of 'filename' is supported by Excel
    """
    return checkxlFileFormat(filename, raiseerror)
