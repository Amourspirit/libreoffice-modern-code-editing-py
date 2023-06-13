"""
Module to demonstrate SourceForge Methods in the Calc Service	
	
Activate	in create_sheet_example
ClearAll	in clear_contents_v1
ClearFormats	in clear_contents_v2
ClearValues	in clear_contents_v3
CopySheet	in copy_sheet_example
CopySheetFromFile	in copy_from_file_example
CopyToCell	in copy_cells_v1
CopyToRange	in copy_cells_v2
DAvg	in calculate_average
DCount	see Davg
DMax	see Davg
DMin	see Davg
DSum	see Davg
Forms	
GetColumnName	see GetValue
GetFormula	see GetValue
GetValue	in mark_invalid
ImportFromCSVFile	in open_csv_file_v1
ImportFromDatabase	see ImportFromCSVFile
InsertSheet	in create_sheet_example
MoveRange	see CopyToRange
MoveSheet	see InsertSheet 
Offset	in create_random_matrix_v1
RemoveSheet	in remove_sheet_example
RenameSheet	see InsertSheet 
SetArray	in create_random_matrix_v2
SetValue	in create_random_matrix_v1
SetCellStyle	in mark_invalid
SetFormula	see SetValue
SortRange
"""

import itertools
from typing import TYPE_CHECKING, Callable, Iterable, Tuple, cast, Any
import uno

if TYPE_CHECKING:
    # anything inside of TYPE_CHECKING block is ignored when script is executed.
    from com.sun.star.sheet import XSpreadsheet
    from com.sun.star.sheet import XSheetCellCursor
    from com.sun.star.sheet import SheetCellRange
    from com.sun.star.frame import XModel
    from com.sun.star.script.provider import XScriptContext

    XSCRIPTCONTEXT = cast(XScriptContext, None)
else:
    XSpreadsheet = object
    XSheetCellCursor = object
    SheetCellRange = object
    XModel = object

from typing import Union
import scriptforge as SF
import random as rnd
from pathlib import Path


# Creates a 6x6 matrix starting at A1
def create_random_matrix_v1(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    for i, j in itertools.product(range(6), range(6)):
        target_cell = doc.Offset("A1", i, j)
        r = rnd.random()
        if r < 0.5:
            doc.SetValue(target_cell, "EVEN")
        else:
            doc.SetValue(target_cell, "ODD")


# Creates a 6x6 matrix starting at A1
# Uses the method setArray to insert values
def create_random_matrix_v2(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    rnd_word = lambda: "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(6)] for _ in range(6)]
    doc.SetArray("A1", values)


# Creates an mxn matrix starting at A1 and asks the desired size
def create_random_matrix_v3(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    bas = SF.CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    for i in range(int(n_rows)):
        for j in range(int(n_cols)):
            target_cell = doc.Offset("A1", i, j)
            r = rnd.random()
            if r < 0.5:
                doc.SetValue(target_cell, "EVEN")
            else:
                doc.SetValue(target_cell, "ODD")


# Creates an mxn matrix starting at A1 and asks the desired size
# Uses the method setArray to insert values
def create_random_matrix_v4(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    bas = SF.CreateScriptService("Basic")
    n_rows = bas.InputBox("Number of rows")
    n_cols = bas.InputBox("Number of columns")
    rnd_word = lambda: "EVEN" if rnd.random() < 0.5 else "ODD"
    values = [[rnd_word() for _ in range(int(n_cols))] for _ in range(int(n_rows))]
    doc.setArray("A1", values)


# Clear region starting at A1
def clear_region_a1(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    cur_sheet = cast(XSpreadsheet, XSCRIPTCONTEXT.getDocument().CurrentController.ActiveSheet)
    cell = cur_sheet.getCellRangeByName("A1")
    cursor: Union[SheetCellRange, XSheetCellCursor] = cur_sheet.createCursorByRange(cell)
    cursor.collapseToCurrentRegion()
    doc.ClearAll(cursor.AbsoluteName)


# Creates a matrix of size 10x8 with random integers between -20 and 100
def create_values_for_example_2(*args: Any) -> None:
    clear_region_a1()
    doc = SF.CreateScriptService("Calc")
    data = [[rnd.randint(-20, 100) for _ in range(8)] for _ in range(10)]
    doc.SetArray("A1", data)


# Marks cells with negative values as INVALID and apply the 'Bad' cell style
def mark_invalid(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    # Gets address of current selection
    cur_selection: str = doc.CurrentSelection
    # Gets address of first cell in the selection
    _ = doc.Offset(cur_selection, 0, 0, 1, 1)
    for i in range(doc.Height(cur_selection)):
        for j in range(doc.Width(cur_selection)):
            cell = doc.Offset(cur_selection, i, j, 1, 1)
            value = doc.GetValue(cell)
            if value < 0:
                doc.SetValue(cell, "INVALID")
                doc.SetCellStyle(cell, "Bad")


# Example of using ClearAll
def clear_contents_v1(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.ClearAll("B2:B7")


# Example of using ClearFormats
def clear_contents_v2(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.ClearFormats("D2:D7")


# Example of using ClearValues
def clear_contents_v3(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.ClearValues("F2:F7")


# Copying to a single cell
def copy_cells_v1(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.CopyToCell("A1:A4", "C1")


# Copying cells into a larger range
def copy_cells_v2(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.CopyToRange("A1:A4", "E1:F6")


# Copies range from an open file
def copy_range_from_file(*args: Any) -> None:
    # Reference to current document (destination)
    doc = SF.CreateScriptService("Calc")
    # Reference to the source document
    svc = SF.CreateScriptService("UI")
    source_doc = cast(SF.SFDocuments.SF_Calc, svc.GetDocument("DataSource.ods"))
    source_range = source_doc.Range("Sheet1.A1:A5")
    # Pastes the contents into the destination
    doc.CopyToCell(source_range, "A1")


# Inserting new sheet
def create_sheet_example(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.InsertSheet("TestSheet", 2)
    doc.Activate("TestSheet")


# Copying an existing sheet
def copy_sheet_example(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.CopySheet("TestSheet", "Copy_TestSheet")


# Removing a sheet
def remove_sheet_example(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    doc.RemoveSheet("Copy_TestSheet")


# Copies sheet from another file (open or closed)
def copy_from_file_example(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    wb = str(Path(res_path, "DataSource.ods"))
    doc.CopySheetFromFile(wb, "Sheet2", "Copy_Sheet2")


# Example using the DAvg method
def calculate_average(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    bas = SF.CreateScriptService("Basic")
    result = doc.DAvg("A1:E1")
    bas.MsgBox("The average is {:.02f}".format(result))


# Open CSV file JobData_v1.csv using default configuration
def open_csv_file_v1(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csv_file = str(Path(res_path, "JobData_v1.csv"))
    doc.ImportFromCSVFile(csv_file, "A1")


# Open CSV file JobData_v2.csv using default configuration
def open_csv_file_v2(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csv_file = str(Path(res_path, "JobData_v2.csv"))
    doc.ImportFromCSVFile(csv_file, "A1")


# Open CSV file using custom configuration
def open_csv_file_v3(*args: Any) -> None:
    doc = SF.CreateScriptService("Calc")
    res_path = get_res_path(doc)
    csv_file = str(Path(res_path, "JobData_v2.csv"))
    filter_option = "59,34,UTF-8,1"
    doc.ImportFromCSVFile(csv_file, "A1", filter_option)


# gets the path to res dir
def get_res_path(doc: SF.SFDocuments.SF_Calc) -> Path:
    cmp: XModel = doc.XComponent
    url = cmp.getURL()
    bas = SF.CreateScriptService("Basic")
    file = bas.ConvertFromUrl(url)
    doc_path = Path(file)
    doc_dir = doc_path.parent
    return doc_dir / "res"


g_exportedScripts = (
    create_random_matrix_v2,
    create_random_matrix_v3,
    create_random_matrix_v1,
    create_random_matrix_v4,
    clear_region_a1,
    mark_invalid,
    create_values_for_example_2,
    clear_contents_v1,
    clear_contents_v2,
    clear_contents_v3,
    copy_cells_v1,
    copy_cells_v2,
    copy_range_from_file,
    create_sheet_example,
    copy_sheet_example,
    remove_sheet_example,
    copy_from_file_example,
    calculate_average,
    open_csv_file_v1,
    open_csv_file_v2,
    open_csv_file_v3,
)

# everything below is only for debugging outside of LibreOffice.


def _ooo_dev_debug(methods: Iterable[Tuple[Callable[[Any], Any], str, tuple]]) -> None:
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
    from ooodev.office.calc import Calc

    def select_used_range(doc: Any, sheet: XSpreadsheet) -> None:
        rng = Calc.find_used_range_obj(sheet=sheet)
        Calc.set_selected_range(doc=doc, sheet=sheet, range_val=rng)

    def clear_selected_range(sheet: XSpreadsheet) -> None:
        rng = Calc.find_used_range_obj(sheet=sheet)
        Calc.clear_cells(sheet=sheet, range_val=rng)

    host = "localhost"
    port = 2002
    # the default ConnectSocket() uses host of localhost and port 2002,
    # but in this case ScriptForge needs to be instructed where to find LibreOffice connection.
    with Lo.Loader(Lo.ConnectSocket(host=host, port=port)):
        fnm = Path(__file__).parent / "lib_o_con_2021.ods"
        doc = Calc.open_doc(fnm=fnm)
        GUI.set_visible(visible=True, doc=doc)
        # give the gui a little time to show up
        Lo.delay(500)
        XSCRIPTCONTEXT = Lo.xscript_context
        SF.ScriptForge(hostname=host, port=port)

        # A breakpoint can be set on the next line to debug the method
        for method in methods:
            meth, name, args = method
            if name:
                sheet = Calc.get_sheet(doc=doc, sheet_name=name)
                Calc.set_active_sheet(doc=doc, sheet=sheet)
                Lo.delay(300)
            else:
                sheet = Calc.get_active_sheet(doc=doc)
            if meth.__name__ in ("mark_invalid", "create_values_for_example_2"):
                select_used_range(doc=doc, sheet=sheet)
            meth(*args)
            # delay to see changes
            Lo.delay(300)
            if meth.__name__ in ("open_csv_file_v1", "open_csv_file_v2", "open_csv_file_v3"):
                clear_selected_range(sheet)
        # close the document
        Lo.close_doc(doc)


def _debug_methods() -> None:
    # these method can be turned on and off for debugging
    # These are mainly for demonstration purposes. Your code may look different.
    methods = (
        (clear_region_a1, "Example1", ()),
        (create_random_matrix_v1, "Example1", ()),
        (clear_region_a1, "Example1", ()),
        (create_random_matrix_v2, "Example1", ()),
        # (create_random_matrix_v3, "Example1", ()), # requires user input
        # (create_random_matrix_v4, "Example1", ()), # requires user input
        (create_values_for_example_2, "Example2", ()),
        (mark_invalid, "Example2", ()),
        (clear_contents_v1, "Example3", ()),
        (clear_contents_v2, "Example3", ()),
        (clear_contents_v3, "Example3", ()),
        (copy_cells_v1, "Example4", ()),
        (copy_cells_v2, "Example4", ()),
        # (copy_range_from_file, "Example4", ()), # requires user input, res/DataSource.ods must be open
        (create_sheet_example, "", ()),  # inserts a sheet
        (copy_sheet_example, "TestSheet", ()),
        (remove_sheet_example, "Copy_TestSheet", ()),
        # (copy_from_file_example, "", ()),
        # (calculate_average, "Example 7", ()),  # requires user input
        (open_csv_file_v1, "Example8", ()),
        (open_csv_file_v2, "Example8", ()),
        (open_csv_file_v3, "Example8", ()),
    )
    _ooo_dev_debug(methods=methods)


if __name__ == "__main__":
    # LibreOffice will never call this.
    #
    # running this script in debug mode, in editor such as Vs Code,
    # will start LibreOffice and run the methods
    _debug_methods()
