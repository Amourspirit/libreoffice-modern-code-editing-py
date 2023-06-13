"""
This is a simple example script to show how to write a script for LibreOffice.

It contains a function that writes "Hello World" into the current document.

This scrips can be run from the Ide or run as a Python Macro from LibreOffice.

The _ooo_dev_debug only works when running in debug mode from the Ide.

The _ooo_dev_debug function starts LibreOffice and runs the method that it is passed.
Then the function closed the document and LibreOffice.

This is useful to debug a script that needs LibreOffice to run with the power of a real IDE.
"""
# coding: utf-8
from __future__ import annotations, unicode_literals
from typing import Any, cast, TYPE_CHECKING, Callable
import uno

if TYPE_CHECKING:
    # Import while design-time to get type hints
    from com.sun.star.text import XSimpleText
else:
    # type hints are not available at runtime
    XSimpleText = object


def hello_world(*args: Any) -> None:
    """
    Generic hello world function.

    Writes Hello World into the current document.
    """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    # cast just to have type support in IDE
    text = cast(XSimpleText, model.Text)
    cursor = text.createTextCursor()
    text.insertString(cursor, "Hello World", False)


g_exportedScripts = (hello_world,)

# everything below is only for debugging outside of LibreOffice.


def _ooo_dev_debug(meth: Callable[[Any], Any], *args) -> None:
    """
    General purpose debug function to run a method in a new document.

    Starts LibreOffice and runs the given method in a new document.

    Args:
        meth (Callable[[Any],Any]): The method to run.
        *args: The arguments to pass to the method.
    """
    # XSCRIPTCONTEXT is not available when running in debug mode from LibreOffice
    # so we need to create it ourselves and get the xscript_context from Lo class.
    global XSCRIPTCONTEXT
    from ooodev.utils.lo import Lo
    from ooodev.utils.gui import GUI
    from ooodev.office.write import Write

    with Lo.Loader(Lo.ConnectSocket()):
        doc = Write.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        # give the gui a little time to show up
        Lo.delay(500)
        XSCRIPTCONTEXT = Lo.xscript_context

        # A breakpoint can be set on the next line to debug the method
        meth(*args)
        # delay to see changes
        Lo.delay(3_000)
        # close the document
        Lo.close_doc(doc)


if __name__ == "__main__":
    # only possible in debug mode
    # LibreOffice will never call this
    _ooo_dev_debug(hello_world)
