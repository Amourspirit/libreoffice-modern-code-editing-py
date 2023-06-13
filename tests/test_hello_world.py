# region imports
import pytest
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from src.hello_world import hello_world

# endregion imports

# This is just a sample test to show how to use pytest with LibreOffice
# For hundreds of examples see: https://github.com/Amourspirit/python_ooo_dev_tools/tree/main/tests

# region run direct

# this is not necessary but allows to run this test directly from python or Editor
if __name__ == "__main__":
    pytest.main([__file__])

# endregion run direct


def test_hello_world(loader, monkeypatch: pytest.MonkeyPatch) -> None:
    # The hello_world.py module has the ability to be run directly and debugged directly.
    # This test is to demonstrate how to test the hello_world.py module using pytest.
    # This is the recommended way to test python code that is intended to be run in LibreOffice.
    if Lo.bridge_connector.headless:
        pytest.skip("Test skipped in headless mode")
    doc = Write.create_doc()
    GUI.set_visible(visible=True, doc=doc)
    Lo.delay(500)
    # Mock the module XSCRIPTCONTEXT variable.
    # the hello_world() method expect the XSCRIPTCONTEXT variable to be set.
    monkeypatch.setattr(hello_world, "XSCRIPTCONTEXT", Lo.xscript_context, raising=False)

    # run the hello_world() method.
    hello_world.hello_world()

    cursor = Write.get_cursor(doc)
    cursor.gotoStart(False)
    cursor.gotoEnd(True)
    doc_text = cursor.getText().getString()
    assert doc_text == "Hello World"

    Lo.close_doc(doc)
