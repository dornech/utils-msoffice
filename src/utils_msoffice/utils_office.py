# Utilities to work with MS Office and other applications from Python via COM interface technology
# general utilities

"""
Module provides various utilities to support COM linking between a Python program and COM application.
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
# ruff: noqa: DTZ005, E115, E501, PLR1714, RUF059, SIM102
#
# disable mypy errors
# - mypy error "'object' has no attribute 'xyz' [attr-defined]" when accessing attributes of
#   dynamically bound wrapped COM object
# mypy: disable-error-code = attr-defined

# fmt: off



from typing import Any, ClassVar
from collections.abc import Callable

import sys
import os.path
# import pathlib
import atexit
import contextlib
import functools
import inspect

import pythoncom
import pywintypes
import win32com.client
import struct

import datetime
import dateutil
import dateutil.tz



# Office ProgIDs / COM class names
COMclass_Access = "Access.Application"
COMclass_Excel = "Excel.Application"



# exception and error handling

# exception class
class ErrorUtilsOffice(BaseException):
    pass


def fix_hresult(hresult: int) -> Any:
    """
    fix_hresult - convert hresult to unsigned long for correct hex representation

    HRESULT values in Windows are typically signed 32-bit integers. This function converts
    a possibly negative signed HRESULT into its unsigned equivalent so that hexadecimal
    formatting (e.g., `0x80004005`) appears correctly.

    Args:
        hresult (int): signed HRESULT value (maybe negative)

    Returns:
        int: unsigned 32-bit representation of the HRESULT
    """
    return struct.unpack("L", struct.pack("l", hresult))[0]



# COM link related routines


def generate_typelib(appCOMclass: str) -> None:
    """
    generate_typelib - generate typelib if not pre-generated as required by COM wrapper (dynamic dispatch does not work)

    Args:
        appCOMclass (str): ProgID / COM class of Microsoft Office host application
    """
    try:
        # get TypeInfo
        typeinfo = win32com.client.Dispatch(appCOMclass)._oleobj_.GetTypeInfo(0)
        # extract the TypeLib and index
        typelib, index = typeinfo.GetContainingTypeLib()
        # get the TypeLib attributes
        typelib_attr = typelib.GetLibAttr()
        # generate module
        win32com.client.gencache.EnsureModule(typelib_attr[0], typelib_attr[1], typelib_attr[3], typelib_attr[3])
        print(f"Successfully generated the makepy module for TypeLibCLSID {typelib_attr[0]} belonging to ProgID / COM class '{appCOMclass}'.")
    except Exception as e:
        print(f"Error during  generation of makepy module for TypeLibCLSID {typelib_attr[0]} belonging to ProgID / COM class '{appCOMclass}': {e}")


def assignCOMapplication(appCOMclass: str, tryStart: bool) -> tuple[object, bool]:
    """
    assignCOMapplication - generalised start routine for Microsoft Office host applications

    Args:
        appCOMclass (str): ProgID / COM class of Microsoft Office host application
        tryStart (bool): flag for start if not already started

    Returns:
        Tuple[object, bool]: application COM object and flag if started by function
    """

    appCOMobj: object = None
    started_app: bool = False

    pythoncom.CoInitialize()

    try:
        appCOMobj = win32com.client.GetActiveObject(Class=appCOMclass)
        if appCOMobj is not None:
            if isinstance(appCOMobj, win32com.client.CDispatch) or hasattr(appCOMobj, "_olerepr_"):
                generate_typelib(appCOMclass)
                appCOMobj = win32com.client.GetActiveObject(Class=appCOMclass)
        # property 'visible' not supported by all Microsoft Office applications
        with contextlib.suppress(Exception):
            appCOMobj.Visible = True
        #  error code  0x800401E2 "Operation unavailable" is equivalent to VBA runtime error 429
        #  https://support.microsoft.com/en-us/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application
    except pythoncom.com_error as ErrorApplicationNotStarted:
        if tryStart:
            try:
                appCOMobj = win32com.client.gencache.EnsureDispatch(appCOMclass)
                appCOMobj.Visible = True
                started_app = True
                # make sure Microsoft Office application is closed if opened by Python program
                atexit.register(startedAppQuit, appCOMobj)
            except pythoncom.com_error as ErrorApplicationNotStartable:
                hresult, msg, exc, arg = ErrorApplicationNotStartable.args
                hresultfixed = fix_hresult(hresult)
                err_msg = f"COM application with ProgID / COM class '{appCOMclass}' could not be started.\nCOM error HRESULT: {hresult} / {hex(hresultfixed)}\nCOM error msg: {msg}"
                raise ErrorUtilsOffice(err_msg)  # noqa B904
        else:
            hresult, msg, exc, arg = ErrorApplicationNotStarted.args
            hresultfixed = fix_hresult(hresult)
            err_msg = f"No active instance of COM application for ProgID/ COM class '{appCOMclass}' found.\nCOM error HRESULT: {hresult} / {hex(hresultfixed)}\nCOM error msg: {msg}"
            raise ErrorUtilsOffice(err_msg)  # noqa B904

    return appCOMobj, started_app

def assign_COMapplication(appCOMclass: str, try_start: bool) -> tuple[object, bool]:
    """
    assign_COMapplication - generalised start routine for Microsoft Office host applications

    Args:
        appCOMclass (str): ProgID / COM class of Microsoft Office host application
        try_start (bool): flag for start if not already started

    Returns:
        Tuple[object, bool]: application COM object and flag if started by function
    """
    return assignCOMapplication(appCOMclass, try_start)


def assignCOMdocument(docfile: str) -> object | None:
    """
    assignCOMdocument - assign Microsoft Office document

    Args:
        docfile (str): filename of Microsoft Office document

    Returns:
        [object or None]: document COM object
    """

    docCOMobj: object = None

    pythoncom.CoInitialize()

    # no path provided as part of name of docfile, only filename -> use execution directory as default
    if docfile == os.path.basename(docfile):
    # if docfile == pathlib.Path(docfile).name:
        docfile = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), docfile)
        # docfile = pathlib.Path(sys.argv[0]).absolute().parent.joinpath(docfile)

    try:
        docCOMobj = win32com.client.GetObject(docfile)
    except pythoncom.com_error as ErrorDocumentNotFound:
        hresult, msg, exc, arg = ErrorDocumentNotFound.args
        hresultfixed = fix_hresult(hresult)
        err_msg = f"Document '{docfile}' not found or not open in COM application.\nCOM error HRESULT: {hresult} / {hex(hresultfixed)}\nCOM error msg: {msg}"
        raise ErrorUtilsOffice(err_msg) # noqa

    # special treatment for Microsoft Access
    try:
        if docCOMobj.Name == docCOMobj.Parent.Name:
            if docfile == docCOMobj.Parent.CurrentProject.FullName:
                return docCOMobj.Parent.CurrentProject  # type: ignore[no-any-return]
            else:
                err_msg = f"Error assigning Microsoft Access Document '{docfile}'."
                raise ErrorUtilsOffice(err_msg)
    except (TypeError, AttributeError):
        pass

    return docCOMobj

def assign_COMdocument(docfile: str) -> object | None:
    """
    assign_COMdocument - assign Microsoft Office document

    Args:
        docfile (str): filename of Microsoft Office document

    Returns:
        [object or None]: document COM object
    """
    return assignCOMdocument(docfile)


def setAppStatus(appCOMobj: object, status: str | bool) -> None:
    """
    setAppStatus - set Microsoft Office application status

    Args:
        appCOMobj (_type_): Microsoft Office application COM object
        status(str, bool): status
    """
    # try:
    #     appCOMobj.StatusBar = status
    # except:
    #     pass
    with contextlib.suppress(Exception):
        appCOMobj.StatusBar = status

def set_app_status(appCOMobj: object, status: str | bool) -> None:
    """
    set_app_status - set Microsoft Office application status

    Args:
        appCOMobj (_type_): Microsoft Office application COM object
        status(sr, bool): status
    """
    setAppStatus(appCOMobj, status)


# take care app is closed if started by Python
# execution might be delayed after exiting Python
def startedAppQuit(appCOMobj) -> None:
    """
    startedAppQuit - quit Microsoft Office application (f. e. used in atexit)

    Args:
        appCOMobj (_type_): Microsoft Office application COM object
    """
    if appCOMobj is not None:
        appCOMobj.StatusBar = ""
        appCOMobj.Quit()
        appCOMobj = None

def quit_started_app(appCOMobj) -> None:
    """
    quit_started_app - quit Microsoft Office application (f. e. used in atexit)

    Args:
        appCOMobj (_type_): Microsoft Office application COM object
    """
    startedAppQuit(appCOMobj)


def get_office_constant(constant: str) -> Any:
    """
    get_office_constant - get Microsoft Office constant from module generated by PyWin32's make.py

    Args:
        constant (str): identifier of desired Microsoft Office constant

    Returns:
        Any: value of constant
    """
    # return win32com.client.constants.__getattr__(constant)
    return getattr(win32com.client.constants, constant)



# utilities for generic wrapper object class and usage

class GenericIterator:
    """
    iterator class to overcome different index base between VBA and Python list comprehension
    """

    def __init__(self, data):  # docsig: disable=SIG102
        self.data = data
        self.index = 0

    def __next__(self):
        self.index += 1
        if self.index < self.data.Count:
            return self.data(self.index)
        else:
            raise StopIteration

    def __iter__(self):
        return self


def get_attrmap(cls: object) -> dict:
    """
    get_attrmap - helper to generate mapping for class camel case method/property identifiers for getattr
    """

    attrmap = {}

    for direntry in dir(cls):
        if direntry[:2] != "__" and direntry[-2:] != "__":
            if hasattr(cls, direntry):
                if callable(getattr(cls, direntry)):
                    direntry_lower = direntry.lower()
                    attrmap[direntry_lower] = direntry

    return attrmap

def get_attrmapCOM(COMobj: object) ->  tuple[dict, dict, dict]:
    """
    get_attrmapCOM - helper to generate mapping for Microsoft Office camel case method/property identifiers for getattr and setattr
    """

    attrmap_get = {}
    attrmap_put = {}
    attrmap_method = {}

    # check if makepy done - _olerepr_ object is late binding and causes problems
    if hasattr(COMobj, "_olerepr_"):
        err_msg = f"makepy not executed for object of type '{COMobj._username_}', late binding not supported for wrapping."
        raise ErrorUtilsOffice(err_msg)
    # determine parent object for get/put properties
    baseobj = COMobj._dispobj_ if hasattr(COMobj, "_dispobj_") else COMobj

    if hasattr(baseobj, "_prop_map_get_"):
        for key in baseobj._prop_map_get_:
            if key[:1] != "_":
                attrmap_get[key.lower()] = key
        attrmap_get = dict(sorted(attrmap_get.items()))
    else:
        err_msg = "'_prop_map_get_' not found in COM object"
        raise ErrorUtilsOffice(err_msg)

    if hasattr(baseobj, "_prop_map_put_"):
        for key in baseobj._prop_map_put_:
            if key[:1] != "_":
                attrmap_put[key.lower()] = key
        attrmap_put = dict(sorted(attrmap_put.items()))
    else:
        err_msg = "'_prop_map_put_' not found in COM object"
        raise ErrorUtilsOffice(err_msg)

    for direntry in dir(baseobj):
        if direntry[0] != "_" and direntry != "CLSID" and direntry != "coclass_clsid" and direntry.lower()[:5] != "dummy":
            if direntry.lower() not in attrmap_get and direntry.lower() not in attrmap_put:
                attrmap_method[direntry.lower()] = direntry
    attrmap_method = dict(sorted(attrmap_method.items()))

    return attrmap_get, attrmap_put, attrmap_method

    return attrmap_get, attrmap_put, attrmap_method


def callwrapper_COMmethod(wrapped_object: object, method: str, wrap_retval: Callable, *args, **kwargs):
    """
    callwrapper - wrapper for COM methods

    The call wrapper for COM methods allows a more 'pythonic' call of Microsoft Office
    COM object methods without duplicating the API. The call wrapper is used by the
    respective wrapper classes.

    The wrapper does
    - map keyword arguments to the keywords derived form the makepy generated object wrapper
    - unwrapping parameters for COM calls
    - applies the return value wrapping for method results

    Example 1: the Excel method Workbook.SaveAs can be called with keyword parameters
    in original writing, in lower case or in snake case wirting:
    - xlWorkbook.SaveAs(SaveChanges=True)
    - xlWorkbook.SaveAs(savechanges=True)
    - xlWorkbook.SaveAs(save_changes=True)

    Example 2: passing wrapped objects works as normal with xlWorkbook being a wrapped
    Excel workbook object and xlWorkbook.worksheets[2] being a wrapped Excel worksheet
    object (done via automated wrapping of return values):
    - xlWorkbook.Worksheets.Add(After=xlWorkbook.worksheets[2], Count=2)
    """

    if len(args) != 0 or len(kwargs) != 0:

        # save signature
        attr_signature = inspect.signature(getattr(wrapped_object, method))

        # unwrap parameters
        args_list = list(args)
        for i, arg in enumerate(args):
            if hasattr(arg, "_msoWrapped"):
                args_list[i] = arg._msoWrapped
            if hasattr(arg, "_xlWrapped"):
                args_list[i] = arg._xlWrapped
        args_new = tuple(args_list)
        for key, value in kwargs.items():
            if hasattr(value, "_msoWrapped"):
                kwargs[key] = value._msoWrapped
            if hasattr(value, "_xlWrapped"):
                kwargs[key] = value._xlWrapped
        # map kwargs
        kwargs_new = {}
        for key, value in kwargs.items():
            found_kwarg = False
            for param in attr_signature.parameters:
                if param.lower() == key.lower().replace("_", ""):
                    kwargs_new[param] = value
                    found_kwarg = True
                    break
            if not found_kwarg:
                err_msg = f"function '{method}' has no parameter '{key}'"
                raise AttributeError(err_msg)

        # bind parameters
        bound_params = attr_signature.bind(*args_new, **kwargs_new)
        bound_params.apply_defaults()

        # execute wrapped method
        retval = getattr(wrapped_object, method)(*bound_params.args, **bound_params.kwargs)
        return wrap_retval(retval)

    else:
        retval = getattr(wrapped_object, method)()
        return wrap_retval(retval)


# wrapper classes for "pythonic" call of object methods

class msoBaseWrapper:
    """
    msoBaseWrapper - generic wrapper for Office objects as pass through for calls to wrapped object

    The base  wrapper allows a more 'pythonic' access to the Microsoft Office
    application API without duplicating the API. Precondition is that the underlying
    application objects are wrapped by wrapper classes.

    To allow recursive resolution for a generic wrapper class for Microsoft Office objects
    not being wrapped by application specific class a general mso wrapper class is defined.
    """

    # _cls_attrmap: dict = {}
    _cls_attrmap_wrapped_get: ClassVar[dict] = {}
    _cls_attrmap_wrapped_put: ClassVar[dict] = {}
    _cls_attrmap_wrapped_method: ClassVar[dict] = {}

    def __init__(self, msoWrapped):  # docsig: disable=SIG102

        # initialize instance - version with attribute mappings determined
        # during every instantiation (different object classes!)
        # however: problems are to be expected if wrapper wraps different
        # application classes at the same time because class attributes are
        # shared across instances -> to avoid see use of class factory functions

        # initialize instance - mapping tables
        self.__class__._cls_attrmap_wrapped_get, self.__class__._cls_attrmap_wrapped_put, self.__class__._cls_attrmap_wrapped_method = get_attrmapCOM(msoWrapped)

        # initialize instance - wrapped object
        self._msoWrapped = msoWrapped

    # pass through for Microsoft Office object model get-attributes to wrapped COM object (properties and methods)
    def __getattr__(self, attr) -> Any:

        # getattr for "pythonized" caller names - only wrapped object reference, private wrapper attributes and
        # wrapped object methods for get
        # NOTE: directly available attributes/methods in wrapper class are caught via __getattribute__
        attr_lower = attr.lower()
        attr_desnaked = attr_lower.replace("_", "")
        if attr == "_msoWrapped" or attr[0:1] == "_":
            retval = getattr(self, attr)
        elif attr_lower in self.__class__._cls_attrmap_wrapped_get:
            retval = getattr(self._msoWrapped, self.__class__._cls_attrmap_wrapped_get[attr_lower])
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_get:
            retval = getattr(self._msoWrapped, self.__class__._cls_attrmap_wrapped_get[attr_desnaked])
        elif attr.lower() in self.__class__._cls_attrmap_wrapped_method:
            retval = functools.partial(callwrapper_COMmethod, self._msoWrapped, self.__class__._cls_attrmap_wrapped_method[attr_lower], self._wrap_retval)
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_method:
            retval = functools.partial(callwrapper_COMmethod, self._msoWrapped, self.__class__._cls_attrmap_wrapped_method[attr_desnaked], self._wrap_retval)
        elif hasattr(self._msoWrapped, attr):
            if attr not in self.__class__._cls_attrmap_wrapped_method.values():
                retval = getattr(self._msoWrapped, attr)
            else:
                retval = functools.partial(callwrapper_COMmethod, self._msoWrapped, attr, self._wrap_retval)
        else:
            err_msg = f"'{self!r}' object has no attribute '{attr}'"
            raise AttributeError(err_msg)

        return self._wrap_retval(retval)

    # internal return object wrapper - must be preceeded by _ !!!
    @staticmethod
    def _wrap_retval(retval: object) -> object:

        # wrap return value for recursion
        if hasattr(retval, "_prop_map_get_"):
            if retval._prop_map_get_.get("Count", False):
                # here not yet wrapped Collection objects are identified
                retval = msoCollectionWrapper(retval)
            elif retval._prop_map_get_.get("Application", False):
                # here other objects are identified
                retval = msoBaseWrapper(retval)
        return retval

    # pass through for Microsoft Office object model set-attributes to wrapped COM object
    def __setattr__(self, attr, value) -> Any:

        # setattr for "pythonized" caller names -  only wrapped object reference, private wrapper attributes and
        # wrapped object methods for get for set
        attr_lower = attr.lower()
        attr_desnaked = attr_lower.replace("_", "")
        if attr == "_msoWrapped" or attr[0:1] == "_":
            super().__setattr__(attr, value)
        elif attr_lower in self.__class__._cls_attrmap_wrapped_put:
            setattr(self._msoWrapped, self.__class__._cls_attrmap_wrapped_put[attr_lower], value)
        elif attr_desnaked in self.__class__._cls_attrmap_wrapped_put:
            setattr(self._msoWrapped, self.__class__._cls_attrmap_wrapped_put[attr_desnaked], value)
        # setattr must consider own class but not call hasattr / __getattr__ (avoid recursion)
        elif attr in self.__class__._cls_attrmap_wrapped_put.values():  # equivalent to hasattr(self._msoWrapped, attr)
            # __setattr__ must not call hasattr / __getattr__ (avoid recursion)
            setattr(self._msoWrapped, attr, value)
        else:
            err_msg = f"'{self!r}' object has no attribute '{attr}'"
            raise AttributeError(err_msg)

def create_msoBaseWrapper(msoWrapped: object):
    """
    create_msoBaseWrapper - class factory function for msoBaseWrapper
    """

    class msoBaseWrapper_created(msoBaseWrapper):
        pass

    return msoBaseWrapper_created(msoWrapped)


class msoCollectionWrapper(msoBaseWrapper):
    """
    msoCollectionWrapper - wrapper for collection objects with pass through for direct calls to wrapped object
    """

    def __call__(self, idx: int | str):

        if isinstance(idx, (int, str)):
            return msoBaseWrapper(self._msoWrapped.Item(idx))
        else:
            return self._msoWrapped

# creator function - required to create/allow multiple class instances
def create_msoCollectionWrapper(msoWrapped: object):
    """
    create_msoCollectionWrapper - class factory function for msoCollectionWrapper
    """

    class msoCollectionWrapper_created(msoCollectionWrapper):
        pass

    return msoCollectionWrapper_created(msoWrapped)



# utilities for callback
# http://exceldevelopmentplatform.blogspot.com/2020/04/vba-calling-python-calling-back-into-vba.html


def ensureDispatch(COMobj):
    """
    ensureDispatch - ensure dispatch (wrapper if PyIDispatch is provided)
    """

    try:
        dispCOMobj = None
        apptypename = str(type(COMobj))
        if apptypename == "<class 'win32com.client.CDispatch'>":
            # this call from GetObject so no need to Dispatch()
            dispCOMobj = COMobj
        elif apptypename == "<class 'PyIDispatch'>":
            # this was passed in from VBA so wrap in Dispatch
            dispCOMobj = win32com.client.Dispatch(COMobj)
        else:
            # other cases just attempt to wrap
            dispCOMobj = win32com.client.Dispatch(COMobj)
        return dispCOMobj
    except Exception as DispatchException:
        if hasattr(DispatchException, "message"):
            return "Error: " + DispatchException.message
        else:
            return "Error: " + str(DispatchException)

def ensure_dispatch(COMobj):
    """
    ensure_dispatch - ensure dispatch (wrapper if PyIDispatch is provided)
    """
    return ensureDispatch(COMobj)


def enhanceErrorMsg(exception: Exception, localsinfo: dict) -> str:
    """
    enhanceErrorMsg - enhance locals (for context of callback)
    """

    # convert dict to string with carriage return
    localsinfostr2 = "\nLocals: {\n"
    for key, value in localsinfo.items():
        try:
            localsinfostr2 += f"'{key}': {value}\n"
        except BaseException:
            localsinfostr2 += f"'{key}': {value.__repr__}\n"
    localsinfostr2 += "}"

    # build error message
    if hasattr(exception, "message"):
        return "Error:" + exception.message + localsinfostr2  # type: ignore[no-any-return]
    else:
        return "Error:" + str(exception) + localsinfostr2

def enhance_errormsg(exception: Exception, localsinfo: dict) -> str:
    """
    enhance_errormsg - enhance locals (for context of callback)
    """
    return enhanceErrorMsg(exception, localsinfo)



# conversion of Python datetime from/to COM timestamp (i.e. pywintypesTime)

# convert Python datetime to COM timestamp
# issue: when calling COM time is automatically assumed as UTC -> adjustment for timezone other than UTC

def cnv_datetime2COMtime(dtval: datetime.datetime | datetime.date, assumeUTC: bool = True) -> pywintypes.Time:
    """
    cnv_datetime2COMtime - convert Python datetime to COM timestamp (COM time always assumed UTC time)

    Args:
        dtval (Union[datetime.datetime or datetime.date]): Python datetime
        assumeUTC(bool, optional): flag for assuming UTC. Defaults to True.

    Returns:
        pywintypes.Time: COM timestamp adjusted by UTC offset
    """
    return cnv_datetime_COMtime(dtval, assumeUTC)

def cnv_datetime_COMtime(dtval: datetime.datetime | datetime.date, assumeUTC: bool = True) -> pywintypes.Time:
    """
    cnv_datetime_COMtime - convert Python datetime to COM timestamp (COM time always assumed UTC time)

    Args:
        dtval (Union[datetime.datetime or datetime.date]): Python datetime
        assumeUTC(bool, optional): flag for assuming UTC. Defaults to True.

    Returns:
        pywintypes.Time: COM timestamp adjusted by UTC offset
    """
    # if isinstance(dtval, datetime.date):
    if not isinstance(dtval, datetime.datetime):
        dtval = datetime.datetime.combine(dtval, datetime.datetime.min.time())
    if assumeUTC:
        dtval += dateutil.tz.tz.tzlocal().utcoffset(datetime.datetime.now())  # type: ignore[operator]
    return pywintypes.Time(dtval)


# convert COM timestamp to Python datetime
# issue: when retrieving datetime values from COM it cannot be compared to Python datetime

def cnv_COMtime2datetime(dtval: pywintypes.Time) -> datetime.datetime:
    """
    cnv_COMtime2datetime - COM timestamp to Python datetime to make them comparable

    Args:
        dtval (pywintypes.Time): COM timestamp

    Returns:
        datetime.datetime: Python datetime
    """
    return cnv_COMtime_datetime(dtval)

def cnv_COMtime_datetime(dtval: pywintypes.Time) -> datetime.datetime:
    """
    cnv_COMtime_datetime - COM timestamp to Python datetime

    Args:
        dtval (pywintypes.Time): COM timestamp

    Returns:
        datetime.datetime: Python datetime
    """
    return datetime.datetime(  # noqa
        year=dtval.year,
        month=dtval.month,
        day=dtval.day,
        hour=dtval.hour,
        minute=dtval.minute,
        second=dtval.second
    )
