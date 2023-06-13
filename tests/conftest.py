# region Imports
import os
import sys
from pathlib import Path
import shutil
import stat
import tempfile
import pytest
from ooodev.utils.lo import Lo as mLo
from ooodev.utils import paths as mPaths
from ooodev.utils.inst.lo.options import Options as LoOptions

# from ooodev.connect import connectors as mConnectors
from ooodev.conn import cache as mCache

# endregion Imports

# comment out next line if want to run LibreOffice in headless mode.
os.environ["NO_HEADLESS"] = "1"


# region Helper Methods
def remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


# endregion Helper Methods

# region fixtures


@pytest.fixture(scope="session")
def tmp_path_session():
    result = Path(tempfile.mkdtemp())  # type: ignore
    yield result
    if os.path.exists(result):
        shutil.rmtree(result, onerror=remove_readonly)


@pytest.fixture(scope="session")
def soffice_path():
    # allow for a little more development flexibility
    # it is also fine to return "" or None from this function

    # return Path("/snap/bin/libreoffice")

    return mPaths.get_soffice_path()


@pytest.fixture(scope="session")
def soffice_env():
    # for snap testing the PYTHONPATH must be set to the virtual environment
    return {}
    # py_pth = mPaths.get_virtual_env_site_packages_path()
    # py_pth += f":{Path.cwd()}"
    # return {"PYTHONPATH": py_pth}


@pytest.fixture(scope="session")
def loader(tmp_path_session, run_headless, soffice_path, soffice_env):
    # This loader is use in testing.
    # It loads LibreOffice and connect to it.
    # This the connection that test use to interact with LibreOffice.
    #
    # for testing with a snap the cache_obj must be omitted.
    # This because the snap is not allowed to write to the real tmp directory.
    loader = mLo.load_office(
        connector=mLo.ConnectPipe(headless=run_headless, soffice=soffice_path, env_vars=soffice_env),
        cache_obj=mCache.Cache(working_dir=tmp_path_session),
        opt=LoOptions(verbose=True),
    )
    # loader = mLo.load_office(
    #     connector=mLo.ConnectSocket(headless=run_headless, soffice=soffice_path, env_vars=soffice_env),
    #     cache_obj=mCache.Cache(working_dir=tmp_path_session),
    #     opt=LoOptions(verbose=True),
    # )
    yield loader
    mLo.close_office()


@pytest.fixture(scope="session")
def run_headless():
    # windows/powershell
    #   $env:NO_HEADLESS='1'; pytest; Remove-Item Env:\NO_HEADLESS
    # linux
    #  NO_HEADLESS="1" pytest
    no_headless = os.environ.get("NO_HEADLESS", "")
    if no_headless == "1":
        return False
    return True


# region Skip Markers for Headless and Os


@pytest.fixture(autouse=True)
def skip_for_headless(request, run_headless: bool):
    # https://stackoverflow.com/questions/28179026/how-to-skip-a-pytest-using-an-external-fixture
    #
    # Also Added:
    # [tool.pytest.ini_options]
    # markers = ["skip_headless: skips a test in headless mode",]
    # see: https://docs.pytest.org/en/stable/how-to/mark.html
    #
    # Usage:
    # @pytest.mark.skip_headless("Requires Dispatch")
    # def test_write(loader, para_text) -> None:
    if run_headless:
        if request.node.get_closest_marker("skip_headless"):
            reason = ""
            try:
                reason = request.node.get_closest_marker("skip_headless").args[0]
            except Exception:
                pass
            if reason:
                pytest.skip(reason)
            else:
                pytest.skip("Skipped in headless mode")


@pytest.fixture(autouse=True)
def skip_not_headless_os(request, run_headless: bool):
    # Usage:
    # @pytest.mark.skip_not_headless_os("linux", "Errors When GUI is present")
    # def test_write(loader, para_text) -> None:

    if not run_headless:
        rq = request.node.get_closest_marker("skip_not_headless_os")
        if rq:
            is_os = sys.platform.startswith(rq.args[0])
            if not is_os:
                return
            reason = ""
            try:
                reason = rq.args[1]
            except Exception:
                pass
            if reason:
                pytest.skip(reason)
            else:
                pytest.skip(f"Skipped in GUI mode on os: {rq.args[0]}")


# endregion Skip Markers for Headless and Os

# endregion fixtures
